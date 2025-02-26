import os
import sys
import ctypes
import subprocess
import logging
import time
import traceback
import json
import locale
import platform
import shutil
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TimeElapsedColumn, TextColumn
from rich.panel import Panel
from rich import box
from concurrent.futures import ThreadPoolExecutor
import argparse
import queue
import psutil
import re
import requests

# Настройка логирования
def setup_logging(log_level):
    logger = logging.getLogger(__name__)
    logger.setLevel(log_level)

    # Логирование в файл
    file_handler = logging.FileHandler("optimization.log", encoding='utf-8')
    file_handler.setLevel(log_level)
    file_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(file_formatter)

    logger.addHandler(file_handler)

    return logger

console = Console()


def print_banner():
    """Выводит баннер скрипта."""
    logger.debug("Вызов print_banner")
    console.print(Panel("[bold yellow]WINDOWS OPTIMIZATION SCRIPT[/bold yellow]",
                        border_style="cyan", expand=False))
    logger.debug("print_banner завершён")


def run_cmd(command, timeout=300):
    """Выполняет команду в командной строке Windows."""
    logger.debug(f"Запуск команды: {command}")
    encoding = 'utf-8'  # Явно указываем кодировку
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding=encoding,
                                errors='replace', timeout=timeout)
        logger.debug(f"Вывод команды: {result.stdout}")
        if result.stderr:
            logger.warning(f"Ошибки команды: {result.stderr}")
        if result.returncode == 0:
            logger.debug(f"Команда выполнена успешно: {command}")
            return True
        else:
            logger.error(f"Ошибка выполнения команды: {command} - Код возврата: {result.returncode}, Ошибка: {result.stderr}")
            console.print(f"  [red]✗ Ошибка:[/red] {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        logger.error(f"Превышено время ожидания команды: {command}")
        console.print(f"  [red]✗ Превышено время ожидания[/red]")
        return False
    except Exception as e:
        logger.error(f"Исключение при выполнении команды: {command} - {str(e)}\n{traceback.format_exc()}")
        console.print(f"  [red]✗ Исключение:[/red] {str(e)}")
        return False


def run_powershell(command, timeout=300):
    """Выполняет команду в PowerShell."""
    logger.debug(f"Запуск PowerShell команды: {command}")
    encoding = 'utf-8'  # Явно указываем кодировку
    try:
        result = subprocess.run(f'powershell -Command "{command}"', shell=True, capture_output=True, text=True,
                                encoding=encoding, errors='replace', timeout=timeout)
        logger.debug(f"Вывод PowerShell: {result.stdout}")
        if result.stderr:
            logger.warning(f"Ошибки PowerShell: {result.stderr}")
        if result.returncode == 0:
            logger.debug(f"PowerShell команда выполнена: {command}")
            return True
        else:
            logger.error(f"Ошибка PowerShell команды: {command} - {result.stderr}")
            console.print(f"  [red]✗ Ошибка:[/red] {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        logger.error(f"Превышено время ожидания PowerShell команды: {command}")
        console.print(f"  [red]✗ Превышено время ожидания[/red]")
        return False
    except Exception as e:
        logger.error(f"Исключение при выполнении PowerShell: {command} - {str(e)}\n{traceback.format_exc()}")
        console.print(f"  [red]✗ Исключение:[/red] {str(e)}")
        return False


def is_admin():
    """Проверяет, запущен ли скрипт с правами администратора."""
    logger.debug("Проверка прав администратора")
    try:
        if os.name != 'nt':
            logger.warning("Проверка прав администратора поддерживается только на Windows.")
            return False
        result = ctypes.windll.shell32.IsUserAnAdmin()
        logger.debug(f"Результат проверки прав: {result}")
        return result
    except Exception as e:
        logger.error(f"Ошибка проверки прав администратора: {str(e)}\n{traceback.format_exc()}")
        return False


def restart_as_admin():
    """Перезапускает скрипт с правами администратора, если их нет."""
    logger.debug("Проверка необходимости перезапуска с правами администратора")
    if not is_admin():
        logger.info("Перезапуск с правами администратора")
        console.print("[red]⚠ Требуются права администратора. Перезапуск...[/red]")
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
        except Exception as e:
            logger.error(f"Ошибка при перезапуске с правами администратора: {str(e)}\n{traceback.format_exc()}")
            console.print(f"[red]✗ Ошибка перезапуска: {str(e)}[/red]")
    logger.debug("Перезапуск не требуется, права уже есть")


def reg_query(key, value):
    """Проверяет значение в реестре Windows."""
    result = subprocess.run(f'reg query "{key}" /v {value}', shell=True, capture_output=True, text=True,
                            encoding='utf-8', errors='replace')
    return result.returncode == 0 and "0x2" in result.stdout  # Предполагаем, что значение должно быть 0x2


def reg_set(key, value, reg_type, data):
    """Устанавливает значение в реестре Windows."""
    return run_cmd(f'reg add "{key}" /v {value} /t {reg_type} /d {data} /f')


def verify_optimization(task_name):
    """Проверяет результат оптимизации."""
    logger.debug(f"Проверка результата для {task_name}")
    if task_name == "Создание точки восстановления системы":
        result = subprocess.run(r"wmic /namespace:\root\default path SystemRestore get LastRestoreId",
                                shell=True, capture_output=True, text=True, encoding='utf-8')
        return result.returncode == 0 and "Optimization Restore Point" in result.stdout
    elif task_name == "Отключение ненужных служб":
        result = subprocess.run(r"sc query DiagTrack", shell=True, capture_output=True, text=True,
                                encoding='utf-8')
        return "STATE" in result.stdout and "STOPPED" in result.stdout
    elif task_name == "Отключение экрана блокировки":
        return reg_query(r"HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Personalization", "NoLockScreen")
    return True


def create_restore_point(requires_confirmation=False):
    """Создаёт точку восстановления системы."""
    logger.debug("Начало create_restore_point")
    console.print("[bold green]➤ Создание точки восстановления системы...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        vss_result = subprocess.run(r"sc query VSS", shell=True, capture_output=True, text=True,
                                    encoding=locale.getpreferredencoding())
        if "STATE" not in vss_result.stdout or "STOPPED" in vss_result.stdout:
            console.print("  [yellow]⚠ Служба Volume Shadow Copy отключена. Точка восстановления не создана.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True

        success = run_cmd(r'wmic /namespace:\root\default path SystemRestore call CreateRestorePoint "Optimization Restore Point", 100, 7')
        progress.update(task, advance=1)

        if success and verify_optimization("Создание точки восстановления системы"):
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def disable_services(requires_confirmation=False):
    """Отключает ненужные службы Windows."""
    logger.debug("Начало disable_services")
    console.print("[bold green]➤ Отключение ненужных служб...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        services = ["DiagTrack", "SysMain", "WSearch", "Fax", "XblGameSave", "XboxNetApiSvc"]
        all_success = True

        for service in services:
            if not run_cmd(f"sc config {service} start= disabled"):
                all_success = False

        progress.update(task, advance=1)

        if all_success and verify_optimization("Отключение ненужных служб"):
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def optimize_disk_performance(requires_confirmation=False):
    """Оптимизирует производительность диска."""
    logger.debug("Начало optimize_disk_performance")
    console.print("[bold green]➤ Оптимизация производительности диска...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        result = subprocess.run(r"fsutil behavior query disablelastaccess", shell=True, capture_output=True,
                                text=True, encoding=locale.getpreferredencoding())
        if "DisableLastAccess = 1" in result.stdout:
            console.print("  [yellow]⚠ Отслеживание последнего доступа уже отключено.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = run_cmd(r"fsutil behavior set disablelastaccess 1")
        progress.update(task, advance=1)
        if success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def disable_visual_effects(requires_confirmation=False):
    """Отключает визуальные эффекты для ускорения системы."""
    logger.debug("Начало disable_visual_effects")
    console.print("[bold green]➤ Отключение визуальных эффектов...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        key = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
        if reg_query(key, "VisualFXSetting"):
            console.print("  [yellow]⚠ Визуальные эффекты уже отключены.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = reg_set(key, "VisualFXSetting", "REG_DWORD", "2")
        progress.update(task, advance=1)
        if success and reg_query(key, "VisualFXSetting"):
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def clean_temp_files(requires_confirmation=False):
    """Очищает временные файлы."""
    logger.debug("Начало clean_temp_files")
    console.print("[bold green]➤ Очистка временных файлов...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    temp_dir = os.path.expandvars(r"%temp%")
    win_temp = r"C:\Windows\Temp"
    all_success = True
    retry_queue = queue.Queue()
    max_attempts = 3  # Максимальное количество попыток удаления

    def attempt_delete(file_path, attempts=0):
        try:
            os.remove(file_path)
            return True
        except PermissionError:
            if attempts < max_attempts:
                logger.warning(f"Файл занят другим процессом, добавление в очередь: {file_path}")
                retry_queue.put((file_path, attempts + 1))
            else:
                logger.error(f"Не удалось удалить файл после {max_attempts} попыток: {file_path}")
            return False
        except Exception as e:
            logger.error(f"Ошибка удаления файла {file_path}: {str(e)}")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=2)

        # Очистка %temp%
        if os.path.exists(temp_dir):
            progress.update(task, description="Очистка %temp%", advance=1)
            try:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        attempt_delete(os.path.join(root, file))
            except Exception as e:
                logger.error(f"Ошибка при очистке %temp%: {str(e)}")
                all_success = False
        else:
            console.print("  [yellow]⚠ Директория %temp% не найдена.[/yellow]")

        # Очистка C:\Windows\Temp
        if os.path.exists(win_temp):
            progress.update(task, description="Очистка C:\\Windows\\Temp", advance=1)
            try:
                for root, dirs, files in os.walk(win_temp):
                    for file in files:
                        attempt_delete(os.path.join(root, file))
            except Exception as e:
                logger.error(f"Ошибка при очистке C:\\Windows\\Temp: {str(e)}")
                all_success = False
        else:
            console.print("  [yellow]⚠ Директория C:\\Windows\\Temp не найдена.[/yellow]")

        # Повторная попытка удаления файлов из очереди
        while not retry_queue.empty():
            file_path, attempts = retry_queue.get()
            if not attempt_delete(file_path, attempts):
                all_success = False

        if all_success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def set_high_performance(requires_confirmation=True):
    """Устанавливает план высокой производительности."""
    logger.debug("Начало set_high_performance")
    console.print("[bold green]➤ Установка высокой производительности...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        result = subprocess.run(r"powercfg /getactivescheme", shell=True, capture_output=True, text=True,
                                encoding=locale.getpreferredencoding())
        if "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c" in result.stdout:
            console.print("  [yellow]⚠ План высокой производительности уже активен.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = run_cmd(r"powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
        progress.update(task, advance=1)
        if success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def optimize_network(requires_confirmation=False):
    """Оптимизирует настройки сети."""
    logger.debug("Начало optimize_network")
    console.print("[bold green]➤ Оптимизация сети...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=2)
        progress.update(task, description="Отключение автотюнинга TCP", advance=1)
        autotune_success = run_cmd(r"netsh int tcp set global autotuninglevel=disabled")
        progress.update(task, description="Отключение ECN", advance=1)
        ecn_success = run_cmd(r"netsh int tcp set global ecncapability=disabled")
        if autotune_success and ecn_success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def optimize_defender(requires_confirmation=False):
    """Снижает нагрузку от Windows Defender."""
    logger.debug("Начало optimize_defender")
    console.print("[bold green]➤ Оптимизация Windows Defender...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        result = subprocess.run(f'powershell -Command "Get-MpPreference | Select-Object -Property DisableRealtimeMonitoring"',
                                shell=True, capture_output=True, text=True, encoding=locale.getpreferredencoding())
        if "True" in result.stdout:
            console.print("  [yellow]⚠ Мониторинг в реальном времени уже отключен.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = run_powershell(r"Set-MpPreference -DisableRealtimeMonitoring $true")
        progress.update(task, advance=1)
        if success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def disable_telemetry(requires_confirmation=True):
    """Отключает телеметрию Microsoft."""
    logger.debug("Начало disable_telemetry")
    console.print("[bold green]➤ Отключение телеметрии...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        key = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
        if reg_query(key, "AllowTelemetry"):
            console.print("  [yellow]⚠ Телеметрия уже отключена.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = reg_set(key, "AllowTelemetry", "REG_DWORD", "0")
        run_cmd(r'schtasks /change /tn "Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" /disable')
        progress.update(task, advance=1)
        if success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def disable_lock_screen(requires_confirmation=True):
    """Отключает экран блокировки."""
    logger.debug("Начало disable_lock_screen")
    console.print("[bold green]➤ Отключение экрана блокировки...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        key = r"HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Personalization"
        if reg_query(key, "NoLockScreen"):
            console.print("  [yellow]⚠ Экран блокировки уже отключён.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = reg_set(key, "NoLockScreen", "REG_DWORD", "1")
        progress.update(task, advance=1)
        if success and verify_optimization("Отключение экрана блокировки"):
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def disable_game_bar(requires_confirmation=True):
    """Отключает игровую панель Xbox."""
    logger.debug("Начало disable_game_bar")
    console.print("[bold green]➤ Отключение панели игр...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=1)
        key = r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\GameDVR"
        if reg_query(key, "AppCaptureEnabled"):
            console.print("  [yellow]⚠ Панель игр уже отключена.[/yellow]")
            console.print("  [green]✔ Завершено (без изменений).[/green]")
            progress.update(task, advance=1)
            return True
        success = reg_set(key, "AppCaptureEnabled", "REG_DWORD", "0")
        progress.update(task, advance=1)
        if success:
            console.print("  [green]✔ Успешно завершено![/green]")
            return True
        else:
            console.print("  [red]✗ Завершено с ошибкой![/red]")
            return False


def clean_browser_cache(requires_confirmation=True):
    """Очищает кэш браузеров."""
    logger.debug("Начало clean_browser_cache")
    console.print("[bold green]➤ Очистка кэша браузеров...[/bold green]")

    if requires_confirmation:
        while True:
            console.print("[yellow]Продолжить? (y/n): [/yellow]", end="")
            confirm = console.input("").strip().lower()
            if confirm in ['y', 'n']:
                break
            console.print("[yellow]⚠ Введите 'y' или 'n'.[/yellow]")
        if confirm != "y":
            console.print("  [red]✗ Отменено пользователем.[/red]")
            return False

    # Проверка, запущены ли браузеры
    browsers = ["chrome.exe", "firefox.exe", "msedge.exe"]
    for browser in browsers:
        if subprocess.run(f"tasklist /FI \"IMAGENAME eq {browser}\"", shell=True, capture_output=True, text=True).stdout.find(browser) != -1:
            console.print(f"[red]✗ {browser} запущен. Закройте браузер перед очисткой кэша.[/red]")
            return False

    user_profile = os.environ["USERPROFILE"]
    paths = [
        os.path.join(user_profile, r"AppData\Local\Google\Chrome\User Data\Default\Cache"),
        os.path.join(user_profile, r"AppData\Roaming\Mozilla\Firefox\Profiles\*\cache2"),
        os.path.join(user_profile, r"AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\INetCache")
    ]
    all_success = True
    any_cache_found = False

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), TimeElapsedColumn(),
                  transient=True, console=console) as progress:
        task = progress.add_task("Выполнение", total=len(paths))

        for i, path in enumerate(paths):
            browser_name = path.split('\\')[-2]
            progress.update(task, description=f"Очистка {browser_name}", completed=i)
            if os.path.exists(path):
                any_cache_found = True
                console.print(f"  [cyan]Обрабатывается {browser_name}...[/cyan]")
                if not run_cmd(fr'del /q /f /s "{path}\*"'):
                    all_success = False
                    console.print(f"  [red]✗ Ошибка очистки {browser_name}.[/red]")
                else:
                    console.print(f"  [green]✔ {browser_name} очищен.[/green]")
            else:
                console.print(f"  [yellow]⚠ Кэш {browser_name} не найден.[/yellow]")
            progress.update(task, advance=1)

        if not any_cache_found:
            console.print("  [yellow]⚠ Ни один кэш браузера не найден.[/yellow]")
        elif all_success:
            console.print("  [green]✔ Все кэши успешно очищены![/green]")
            return True
        else:
            console.print("  [red]✗ Очистка завершена с ошибками![/red]")
            return False

    return all_success


def save_selection(selected_nums):
    """Сохраняет выбор пользователя в файл."""
    logger.debug(f"Сохранение выбора: {selected_nums}")
    with open("selection.json", "w", encoding="utf-8") as f:
        json.dump(selected_nums, f)
    console.print("[green]✔ Выбор сохранён[/green]")


def load_selection():
    """Загружает выбор пользователя из файла."""
    logger.debug("Загрузка выбора")
    try:
        with open("selection.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        console.print("[yellow]⚠ Файл выбора не найден. Убедитесь, что файл selection.json существует в текущей директории.[/yellow]")
        return None
    except json.JSONDecodeError:
        console.print("[red]✗ Ошибка чтения файла выбора. Проверьте формат файла selection.json.[/red]")
        logger.error("Ошибка чтения файла выбора: неверный формат JSON.")
        return None
    except Exception as e:
        logger.error(f"Ошибка загрузки выбора: {str(e)}\n{traceback.format_exc()}")
        console.print(f"[red]✗ Неизвестная ошибка при загрузке выбора: {str(e)}[/red]")
        return None


def get_user_selection():
    """Получает выбор оптимизаций от пользователя."""
    logger.debug("Начало get_user_selection")
    optimizations = {
        1: ("Создание точки восстановления системы", create_restore_point, False,
            "Создаёт точку восстановления для безопасности.", "Низкое", "Простая"),
        2: ("Отключение ненужных служб", disable_services, False, "Уменьшает нагрузку, отключая службы.", "Среднее", "Средняя"),
        3: ("Оптимизация производительности диска", optimize_disk_performance, False, "Снижает запись данных.", "Среднее", "Простая"),
        4: ("Отключение визуальных эффектов", disable_visual_effects, False, "Ускоряет систему.", "Низкое", "Простая"),
        5: ("Очистка временных файлов", clean_temp_files, False, "Освобождает место.", "Низкое", "Простая"),
        6: ("Установка высокой производительности", set_high_performance, True, "Максимизирует производительность.", "Высокое", "Простая"),
        7: ("Оптимизация сети", optimize_network, False, "Улучшает сеть через TCP.", "Среднее", "Средняя"),
        8: ("Оптимизация Windows Defender", optimize_defender, False, "Снижает нагрузку Defender.", "Среднее", "Средняя"),
        9: ("Отключение телеметрии", disable_telemetry, True, "Прекращает сбор данных.", "Высокое", "Средняя"),
        11: ("Отключение экрана блокировки", disable_lock_screen, True, "Ускоряет вход.", "Низкое", "Простая"),
        14: ("Отключение панели игр", disable_game_bar, True, "Отключает Xbox панель.", "Низкое", "Простая"),
        15: ("Очистка кэша браузеров", clean_browser_cache, True, "Очищает кэш браузеров.", "Среднее", "Средняя"),
    }

    def print_optimizations():
        console.print("[bold cyan]Доступные оптимизации:[/bold cyan]")
        table = Table(show_header=True, header_style="bold magenta", box=box.SIMPLE)
        table.add_column("№", style="cyan", justify="center", width=3)
        table.add_column("Название", style="magenta")
        table.add_column("Описание", style="green")
        table.add_column("Влияние", style="white")
        table.add_column("Сложность", style="white")
        table.add_column("Подтверждение", style="yellow", justify="center")

        for num, (name, _, requires_confirmation, desc, impact, complexity) in optimizations.items():
            impact_colored = f"[green]{impact}[/green]" if impact == "Низкое" else f"[yellow]{impact}[/yellow]" if impact == "Среднее" else f"[red]{impact}[/red]"
            table.add_row(f"{num:02d}", name, desc, impact_colored, complexity, "Да" if requires_confirmation else "Нет")

        console.print(table)
        console.print("[cyan]Введите номера через запятую (например, 01, 02), 'all', 'save', 'load', 'stop':[/cyan]")

    selected = set()
    while True:
        print_optimizations()
        try:
            user_input = console.input("[yellow]Ваш выбор: [/yellow]").strip().lower()
            logger.debug(f"Пользователь ввёл: {user_input}")
        except Exception:
            console.print("[red]✗ Ошибка ввода. Выход.[/red]")
            return None

        if user_input == "stop":
            return None
        elif user_input == "all":
            selected = set(optimizations.keys())
        elif user_input == "save" and selected:
            save_selection(sorted(selected))
        elif user_input == "load":
            loaded = load_selection()
            if loaded:
                selected = set(loaded)
        else:
            try:
                numbers = [int(n.strip()) for n in user_input.split(",") if n.strip().isdigit()]
                selected.update(n for n in numbers if n in optimizations)
            except ValueError:
                console.print("[red]✗ Неверный формат. Пожалуйста, введите корректные номера оптимизаций.[/red]")
                continue

        if selected and user_input not in ["save", "load"]:
            while True:
                console.print("[yellow]Подтвердить выбор? (y/n): [/yellow]", end="")
                confirm = console.input("").strip().lower()
                if confirm in ['y', 'n']:
                    break
                console.print("[yellow]✗ Введите 'y' или 'n'.[/yellow]")
            if confirm == "y":
                save_selection(sorted(selected))
                return [optimizations[num] for num in sorted(selected)]


def print_system_info():
    """Выводит информацию о системе."""
    logger.debug("Вывод информации о системе")
    console.print("[bold cyan]Информация о системе:[/bold cyan]")
    
    # Информация о процессоре
    cpu_usage = psutil.cpu_percent(interval=1)
    console.print(f"[green]Загрузка процессора:[/green] {cpu_usage}%")

    # Информация о дисках
    partitions = psutil.disk_partitions()
    for partition in partitions:
        try:
            usage = psutil.disk_usage(partition.mountpoint)
            console.print(f"[green]Диск {partition.device}:[/green] {usage.free / (1024**3):.2f} GB свободно из {usage.total / (1024**3):.2f} GB")
        except PermissionError:
            logger.warning(f"Нет доступа к диску {partition.device}")

    # Информация о памяти
    memory = psutil.virtual_memory()
    console.print(f"[green]Свободная память:[/green] {memory.available / (1024**3):.2f} GB из {memory.total / (1024**3):.2f} GB")

    # Информация о сети
    net_io = psutil.net_io_counters()
    console.print(f"[green]Сетевые данные:[/green]")
    console.print(f"  [green]Отправлено данных:[/green] {net_io.bytes_sent / (1024**2):.2f} MB (с момента последней перезагрузки)")
    console.print(f"  [green]Получено данных:[/green] {net_io.bytes_recv / (1024**2):.2f} MB (с момента последней перезагрузки)")

    # Проверка пинга
    ping_result = subprocess.run("ping -n 1 google.com", shell=True, capture_output=True, text=True)
    if ping_result.returncode == 0:
        last_line = ping_result.stdout.strip().split('\n')[-1]
        numbers = re.findall(r'\d+', last_line)
        if numbers:
            average_ping = sum(map(int, numbers)) / len(numbers)
            console.print(f"[green]Средний пинг:[/green] {average_ping:.2f} ms")
    else:
        console.print("[red]✗ Не удалось выполнить пинг.[/red]")


def check_for_updates():
    """Проверяет наличие обновлений для скрипта."""
    logger.debug("Проверка обновлений")
    current_version = "0.9.0"
    latest_version = get_latest_version()  # Получите последнюю версию из удаленного источника
    if current_version < latest_version:
        console.print("[yellow]Доступна новая версия скрипта. Пожалуйста, обновите.[/yellow]")
        if console.input("[yellow]Хотите обновить сейчас? (y/n): [/yellow]").strip().lower() == 'y':
            update_script()
    else:
        console.print("[green]У вас установлена последняя версия скрипта.[/green]")


def get_latest_version():
    """Получает последнюю версию скрипта из удаленного источника."""
    # Пример получения версии (замените на реальную логику)
    return "1.0.1"


def update_script():
    """Обновляет скрипт до последней версии."""
    logger.debug("Обновление скрипта")
    try:
        url = "https://example.com/latest_script.py"  # Убедитесь, что это реальный URL
        response = requests.get(url)
        if response.status_code == 200:
            script_path = os.path.abspath(__file__)
            with open(script_path, 'wb') as f:
                f.write(response.content)
            console.print("[green]✔ Скрипт успешно обновлен. Перезапустите программу.[/green]")
        else:
            console.print("[red]✗ Не удалось загрузить обновление.[/red]")
    except Exception as e:
        logger.error(f"Ошибка обновления скрипта: {str(e)}")
        console.print("[red]✗ Ошибка обновления скрипта.[/red]")


def backup_settings():
    """Создает бэкап текущих настроек."""
    logger.debug("Создание бэкапа настроек")
    try:
        if os.path.exists("settings.json"):
            shutil.copy("settings.json", "settings_backup.json")
            console.print("[green]✔ Бэкап настроек успешно создан.[/green]")
        else:
            console.print("[yellow]⚠ Файл settings.json не найден. Бэкап не создан.[/yellow]")
    except Exception as e:
        logger.error(f"Ошибка создания бэкапа: {str(e)}")
        console.print("[red]✗ Ошибка создания бэкапа.[/red]")


def apply_all_optimizations():
    """Применяет все выбранные оптимизации."""
    logger.debug("Начало apply_all_optimizations")
    print_banner()
    print_system_info()  # Выводим информацию о системе
    check_for_updates()  # Проверка обновлений
    backup_settings()  # Создание бэкапа настроек
    restart_as_admin()
    if not check_compatibility() or not check_dependencies():
        console.print("[red]➤ Продолжение невозможно.[/red]")
        return

    selected_optimizations = get_user_selection()
    if not selected_optimizations:
        console.print("[red]➤ Ничего не выбрано.[/red]")
        return

    console.print("\n[bold yellow]Вы выбрали:[/bold yellow]")
    for i, (name, *_) in enumerate(selected_optimizations, 1):
        console.print(f"  [cyan]{i:02d}.[/cyan] {name}")

    console.print("\n[bold cyan]➤ Запуск оптимизаций...[/bold cyan]")
    all_success = True

    # Разделяем задачи на те, что требуют подтверждения, и те, что нет
    confirmation_required = [opt for opt in selected_optimizations if opt[2]]
    no_confirmation = [opt for opt in selected_optimizations if not opt[2]]

    # Выполняем задачи без подтверждения последовательно
    for opt in no_confirmation:
        result = opt[1](opt[2])
        if result is None:
            result = False
        all_success &= result

    # Выполняем задачи с подтверждением последовательно
    for opt in confirmation_required:
        result = opt[1](opt[2])
        if result is None:
            result = False
        all_success &= result

    if all_success:
        console.print(Panel("[bold green]✔ Все настройки применены!\n[yellow]Перезагрузите компьютер.[/yellow]",
                            border_style="cyan", expand=False))
    else:
        console.print(Panel("[bold yellow]⚠ Некоторые оптимизации завершились с ошибками или были отменены.\n[yellow]Перезагрузите компьютер для полного эффекта.[/yellow]",
                            border_style="cyan", expand=False))

    console.input("[cyan]Нажмите Enter для выхода...[/cyan]")


def check_compatibility():
    """Проверяет совместимость с версией Windows."""
    logger.debug("Проверка совместимости")
    version = platform.win32_ver()[1]
    if version < "10":
        console.print("[yellow]⚠ Скрипт протестирован только на Windows 10 и выше.[/yellow]")
        return False
    # Проверка наличия PowerShell
    if shutil.which("powershell") is None:
        console.print("[red]✗ PowerShell не найден. Установите PowerShell для продолжения.[/red]")
        return False
    return True


def check_dependencies():
    """Проверяет наличие необходимых утилит."""
    logger.debug("Проверка зависимостей")
    dependencies = ["sc", "dism", "powershell"]
    missing = [dep for dep in dependencies if shutil.which(dep) is None]
    if missing:
        console.print(f"[red]✗ Отсутствуют: {', '.join(missing)}.[/red]")
        console.print("[yellow]Убедитесь, что все необходимые утилиты установлены и доступны в PATH.[/yellow]")
        console.print("[yellow]Для установки PowerShell, посетите: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-windows[/yellow]")
        return False
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Windows Optimization Script")
    parser.add_argument('--log-level', default='DEBUG', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        help='Уровень логирования (по умолчанию: DEBUG)')
    args = parser.parse_args()

    logger = setup_logging(getattr(logging, args.log_level))
    logger.debug("Запуск программы")

    try:
        apply_all_optimizations()
    except Exception as e:
        logger.error(f"Критическая ошибка: {str(e)}\n{traceback.format_exc()}")
        console.print(f"[red]✗ Ошибка программы: {str(e)}[/red]")
