import time
import subprocess
import sys

import pythoncom
import win32com.client

import win32serviceutil
import win32service
import win32event
import servicemanager
import win32api
import win32con


def is_bitlocker_protected_and_unlocked(drive_letter: str) -> bool:
    """
    Returns True if BitLocker is enabled (Protection On) and volume is not locked.
    drive_letter must be like 'E:' or 'E:\\'.
    """
    if not drive_letter.endswith(":"):
        drive_letter = drive_letter[0] + ":"

    try:
        # Query BitLocker status via manage-bde
        output = subprocess.check_output(
            ["manage-bde", "-status", drive_letter],
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except Exception as exc:
        servicemanager.LogInfoMsg(
            f"UsbBitLockerGuard: manage-bde failed for {drive_letter}: {exc}"
        )
        return False

    protection_on = False
    locked = False

    for line in output.splitlines():
        line = line.strip()
        if line.startswith("Protection Status"):
            # Example: "Protection Status: Protection On"
            if "On" in line:
                protection_on = True
        if line.startswith("Lock Status"):
            # Example: "Lock Status: Unlocked"
            if "Locked" in line:
                locked = True

    return protection_on and not locked


def eject_drive(drive_letter: str):
    """
    Tries to make the removable drive unusable for the user by logically
    ejecting it from Windows. This does NOT permanently change the disk;
    unplugging and re‑plugging it should make it usable again (if the
    service is not running).

    Implementation: send an 'Eject' verb via Shell.Application.
    This is best‑effort and depends on the device and Windows, but it
    avoids persistent side‑effects like 'mountvol /p'.
    """
    if not drive_letter.endswith(":"):
        drive_letter = drive_letter[0] + ":"

    ps_script = (
        "$shell = New-Object -ComObject Shell.Application; "
        "$folder = $shell.NameSpace(17); "  # 17 = My Computer
        f"$item = $folder.ParseName('{drive_letter}'); "
        "if ($item) { $item.InvokeVerb('Eject') }"
    )

    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                ps_script,
            ],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        servicemanager.LogInfoMsg(
            f"UsbBitLockerGuard: Eject command sent for {drive_letter}"
        )
    except Exception as exc:
        servicemanager.LogInfoMsg(
            f"UsbBitLockerGuard: Failed to eject {drive_letter}: {exc}"
        )


def show_alarm(drive_letter: str, msg: str):
    """
    Tries to show a MessageBox; if not possible (service session 0),
    logs to event log instead.
    """
    full_msg = f"USB drive {drive_letter} is NOT BitLocker protected.\n\n{msg}"
    try:
        # This may not be visible when running as a real service on modern Windows,
        # but is useful in 'debug' mode.
        win32api.MessageBox(
            0,
            full_msg,
            "USB BitLocker Guard",
            win32con.MB_ICONWARNING | win32con.MB_OK,
        )
    except Exception:
        servicemanager.LogInfoMsg(f"UsbBitLockerGuard ALARM: {full_msg}")


class UsbBitLockerGuardService(win32serviceutil.ServiceFramework):
    _svc_name_ = "UsbBitLockerGuard"
    _svc_display_name_ = "USB BitLocker Guard"
    _svc_description_ = "Blocks non-BitLocker USB drives by auto-ejecting them."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        # Event used to signal service stop.
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self._stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self._stop_requested = True
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("UsbBitLockerGuard service starting.")
        self.main()
        servicemanager.LogInfoMsg("UsbBitLockerGuard service stopped.")

    def main(self):
        # Initialize COM for this thread
        pythoncom.CoInitialize()

        # Connect to WMI
        c = win32com.client.GetObject("winmgmts:")

        # Watch for new volumes: DriveType=2 (removable) and has drive letter
        query = (
            "SELECT * FROM __InstanceCreationEvent WITHIN 2 "
            "WHERE TargetInstance ISA 'Win32_Volume' "
            "AND TargetInstance.DriveType = 2 "
            "AND TargetInstance.DriveLetter IS NOT NULL"
        )

        watcher = c.ExecNotificationQuery(query)

        while not self._stop_requested:
            try:
                # Wait up to 2 seconds for an event
                event = watcher.NextEvent(2000)
            except pythoncom.com_error:
                # Timeout or transient error, just loop again
                event = None

            # Check if stop was requested during wait
            if self._stop_requested:
                break

            if event is None:
                # No event within timeout
                continue

            try:
                vol = event.TargetInstance
                drive_letter = vol.DriveLetter  # e.g. "E:"
            except Exception:
                continue

            if not drive_letter:
                continue

            servicemanager.LogInfoMsg(
                f"UsbBitLockerGuard: Detected new removable volume {drive_letter}"
            )

            # Small delay to allow Windows to fully mount the volume
            time.sleep(2.0)

            if is_bitlocker_protected_and_unlocked(drive_letter):
                servicemanager.LogInfoMsg(
                    f"UsbBitLockerGuard: {drive_letter} is BitLocker protected and unlocked. Allowed."
                )
            else:
                servicemanager.LogInfoMsg(
                    f"UsbBitLockerGuard: {drive_letter} is NOT BitLocker protected. Ejecting."
                )
                eject_drive(drive_letter)
                show_alarm(
                    drive_letter,
                    "The drive has been ejected automatically because it is not BitLocker-encrypted.",
                )

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Standard pywin32 service entry point:
    #   python usb_bitlocker_service.py install
    #   python usb_bitlocker_service.py start
    #   python usb_bitlocker_service.py stop
    #   python usb_bitlocker_service.py remove
    #   python usb_bitlocker_service.py debug   (for console testing)
    win32serviceutil.HandleCommandLine(UsbBitLockerGuardService)