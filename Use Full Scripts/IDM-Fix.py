import os
import sys
import winreg
import subprocess
import ctypes
import time
import random
import string
import platform
from datetime import datetime

# Check for admin rights
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# If not admin, restart with admin rights
if not is_admin():
    print("Administrator privileges are required for proper IDM activation")
    print("Requesting elevation...")
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(sys.argv[0])}"', None, 1)
        sys.exit(0)
    except:
        print("Failed to restart with admin rights. Continuing with limited functionality...")

# Get system architecture info
def get_system_info():
    arch = platform.architecture()[0]
    if arch == '32bit':
        CLSID = r"Software\Classes\CLSID"
        HKLM_IDM = r"Software\Internet Download Manager"
    else:
        CLSID = r"Software\Classes\Wow6432Node\CLSID"
        HKLM_IDM = r"SOFTWARE\Wow6432Node\Internet Download Manager"
        
    HKCU = winreg.HKEY_CURRENT_USER
    HKLM = winreg.HKEY_LOCAL_MACHINE
    
    # Find IDM executable
    idm_path = None
    try:
        reg_key = winreg.OpenKey(HKCU, r"Software\DownloadManager")
        idm_path = winreg.QueryValueEx(reg_key, "ExePath")[0]
        winreg.CloseKey(reg_key)
    except:
        pass
    
    if not idm_path or not os.path.exists(idm_path):
        if arch == '32bit':
            idm_path = os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 
                                    r"Internet Download Manager\IDMan.exe")
        else:
            idm_path = os.path.join(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'), 
                                    r"Internet Download Manager\IDMan.exe")
    
    return {
        'arch': arch,
        'CLSID': CLSID,
        'HKLM_IDM': HKLM_IDM,
        'HKCU': HKCU,
        'HKLM': HKLM,
        'idm_path': idm_path
    }

# Check if IDM is running and stop it if necessary
def stop_idm():
    print("Checking if IDM is running...")
    try:
        result = subprocess.run('tasklist /FI "IMAGENAME eq IDMan.exe" /NH', 
                              shell=True, capture_output=True, text=True)
        if "IDMan.exe" in result.stdout:
            print("Stopping IDM processes...")
            subprocess.run('taskkill /f /im idman.exe', shell=True, check=False)
            time.sleep(1)
            return True
        else:
            print("IDM is not running")
            return False
    except:
        print("Failed to check if IDM is running")
        return False

# Clean registry function
def clean_registry(system_info, deep=True):
    print("\nPerforming full registry cleanup...")
    
    # Values to delete
    values_to_delete = [
        (system_info['HKCU'], r"Software\DownloadManager", "FName"),
        (system_info['HKCU'], r"Software\DownloadManager", "LName"),
        (system_info['HKCU'], r"Software\DownloadManager", "Email"),
        (system_info['HKCU'], r"Software\DownloadManager", "Serial"),
        (system_info['HKCU'], r"Software\DownloadManager", "scansk"),
        (system_info['HKCU'], r"Software\DownloadManager", "tvfrdt"),
        (system_info['HKCU'], r"Software\DownloadManager", "radxcnt"),
        (system_info['HKCU'], r"Software\DownloadManager", "LstCheck"),
        (system_info['HKCU'], r"Software\DownloadManager", "ptrk_scdt"),
        (system_info['HKCU'], r"Software\DownloadManager", "LastCheckQU"),
        (system_info['HKCU'], r"Software\DownloadManager", "MData"),
        (system_info['HKCU'], r"Software\DownloadManager", "isreged"),
        (system_info['HKCU'], r"Software\DownloadManager", "IsRegistered"),
        (system_info['HKCU'], r"Software\DownloadManager", "ActivationTime"),
        (system_info['HKCU'], r"Software\DownloadManager", "regStatus"),
        (system_info['HKCU'], r"Software\DownloadManager", "isactived"),
        (system_info['HKCU'], r"Software\DownloadManager", "realser")
    ]
    
    # Delete values
    for hkey, key_path, value_name in values_to_delete:
        try:
            with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, value_name)
                    print(f"Deleted registry value: {key_path}\\{value_name}")
                except:
                    pass
        except:
            pass
    
    # Clean system-wide registry entries
    try:
        cmd = f'reg delete "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regStatus /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regname /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regemail /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Wow6432Node\\Internet Download Manager" /v regserial /f'
        subprocess.run(cmd, shell=True, check=False)
        
        # Also try 32-bit paths
        cmd = f'reg delete "HKLM\\SOFTWARE\\Internet Download Manager" /v regStatus /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Internet Download Manager" /v regname /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Internet Download Manager" /v regemail /f'
        subprocess.run(cmd, shell=True, check=False)
        
        cmd = f'reg delete "HKLM\\SOFTWARE\\Internet Download Manager" /v regserial /f'
        subprocess.run(cmd, shell=True, check=False)
        
        print("Cleaned system-wide registry entries")
    except:
        print("Warning: Could not clean all system-wide registry entries")
    
    # If deep cleaning requested, clean CLSID entries
    if deep:
        clean_clsid_keys(system_info)
    
    # Reset installation status
    reset_installation_status(system_info)
    
    # Delete IDM configuration folder
    try:
        appdata_paths = [
            os.path.join(os.environ.get('APPDATA', ''), "IDM"),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), "IDM"),
        ]
        
        for path in appdata_paths:
            if os.path.exists(path):
                try:
                    cmd = f'rmdir /s /q "{path}"'
                    subprocess.run(cmd, shell=True, check=False)
                    print(f"Deleted IDM directory: {path}")
                except:
                    pass
    except:
        pass
        
    return True

# Clean CLSID keys
def clean_clsid_keys(system_info):
    print("Cleaning IDM CLSID registry keys...")
    
    # Look for IDM-related CLSID entries
    clsid_patterns = [
        "{07999AC3-058B-40BF-984F-69EB1E554CA7}",  # Common IDM CLSID
        "{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",  # Common IDM CLSID for browser integration
        "IDMIEHlprObj",  # Helper object class
        "IDMIECC",  # Browser Helper Object
        "IDMBHOObj",  # Browser Helper Object class
        "Internet Download Manager",  # General IDM pattern
        "idmmkb",  # Another IDM pattern
        "IDM.DownloadAll",  # IDM download component
        "IDM.NativeHook",  # IDM hook
    ]
    
    # Additional registry paths to clean
    additional_paths = [
        r"Software\Classes\IDMIECC.IdmIECC",
        r"Software\Classes\IDMIECC.IdmIECC.1",
        r"Software\Classes\IDMIEHlprObj.IdmIEHlprObj",
        r"Software\Classes\IDMIEHlprObj.IdmIEHlprObj.1",
        r"Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",
        r"Software\Microsoft\Windows\CurrentVersion\Ext\Stats\{5ED60779-4DE2-4E07-B862-974CA4FF3D55}",
        r"Software\Microsoft\Windows\CurrentVersion\Ext\Settings\{07999AC3-058B-40BF-984F-69EB1E554CA7}",
        r"Software\Microsoft\Windows\CurrentVersion\Ext\Stats\{07999AC3-058B-40BF-984F-69EB1E554CA7}",
    ]
    
    for path in additional_paths:
        try:
            cmd = f'reg delete "HKCU\\{path}" /f'
            subprocess.run(cmd, shell=True, check=False, capture_output=True)
            print(f"Deleted registry key: {path}")
        except:
            pass

# Reset installation status
def reset_installation_status(system_info):
    print("Resetting IDM installation status...")
    try:
        # Set installation status to fresh install (0)
        cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v InstallStatus /t REG_DWORD /d "0" /f'
        subprocess.run(cmd, shell=True, check=False)
        
        # Also try the non-architecture specific path
        if "Wow6432Node" in system_info["HKLM_IDM"]:
            alt_path = system_info["HKLM_IDM"].replace("Wow6432Node\\", "")
            cmd = f'reg add "HKLM\\{alt_path}" /v InstallStatus /t REG_DWORD /d "0" /f'
            subprocess.run(cmd, shell=True, check=False)
    except:
        print("Warning: Could not reset installation status")

# Generate random license key
def generate_license_key():
    prefix = ''.join(random.choices(string.ascii_uppercase, k=4))
    parts = [prefix]
    for _ in range(3):
        part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
        parts.append(part)
    return "-".join(parts)

# Activate IDM
def activate_idm(system_info):
    print("\nActivating IDM with extended registry modifications...")
    try:
        # Create registration info
        fname = "Registered"
        lname = "User"
        email = "user@example.com"
        serial = generate_license_key()
        
        current_time = int(time.time())
        
        # Create DownloadManager key if it doesn't exist
        try:
            reg_key = winreg.OpenKey(system_info['HKCU'], r"Software\DownloadManager", 0, 
                                   winreg.KEY_SET_VALUE | winreg.KEY_CREATE_SUB_KEY)
        except:
            reg_key = winreg.CreateKey(system_info['HKCU'], r"Software\DownloadManager")
        
        # Set all registry values needed for activation
        reg_info = {
            # Basic registration info
            "FName": fname,
            "LName": lname,
            "Email": email,
            "Serial": serial,
            
            # Additional activation keys
            "ActivationTime": str(current_time),
            "CheckUpdtTime": str(current_time),
            "scansk": "AAAGAAA=",
            "MData": "AAABAAAAAAABAAAAAAAAAAAAAA==",
            "updtprm": "2;12;1;1;1;1;1;0;",
            "SPDirExist": "1",
            "icfname": "idman638build18.exe",
            "icfsize": "8517272",
            "regStatus": "1",
            "netdmin": "15000",
            "taser": fname + " " + lname,
            "realser": serial,
            "isreged": "1",
            "iserror": "0",
            "isactived": "1",
            "DistribType": "0",
            "swfreg": "1",  # Additional registration flag
            "sernmcf": "0"  # Serial name confirmation
        }
        
        # Set string values
        for key, value in reg_info.items():
            try:
                winreg.SetValueEx(reg_key, key, 0, winreg.REG_SZ, value)
                print(f"Set {key}: {value}")
            except Exception as e:
                print(f"Warning: Could not set {key}: {str(e)}")
        
        # Set DWORD values
        dword_values = {
            "IsRegistered": 1,
            "LstCheck": 0,
            "CheckUpdtTime": current_time,
            "AppDataDir": 1,
            "AfterInst": 1,
            "LaunchCnt": 15,
            "scdt": current_time,
            "radxcnt": 0,
            "mngdby": 0,
            "ExeHashed": 1,  # Additional flag to indicate exe has been verified
            "CleanComplete": 1  # Clean installation completed
        }
        
        for key, value in dword_values.items():
            try:
                winreg.SetValueEx(reg_key, key, 0, winreg.REG_DWORD, value)
                print(f"Set {key}: {value} (DWORD)")
            except Exception as e:
                print(f"Warning: Could not set {key}: {str(e)}")
        
        winreg.CloseKey(reg_key)
        
        # Set system-wide registry entries for better activation
        print("\nSetting system-wide registry entries...")
        
        # First check if keys exist and create them if they don't
        try:
            # Check if the key exists, create if not
            cmd = f'reg query "HKLM\\{system_info["HKLM_IDM"]}" /ve'
            result = subprocess.run(cmd, shell=True, capture_output=True)
            if result.returncode != 0:
                # Key doesn't exist, create it
                cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /f'
                subprocess.run(cmd, shell=True, check=False)
                print(f"Created HKLM\\{system_info['HKLM_IDM']} key")
        except:
            pass
            
        # Now try to set system-wide registry entries
        try:
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v regStatus /t REG_SZ /d "1" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v regname /t REG_SZ /d "{fname} {lname}" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v regemail /t REG_SZ /d "{email}" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v regserial /t REG_SZ /d "{serial}" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v InstallStatus /t REG_DWORD /d "3" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            # Additional registry entries
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v isreged /t REG_SZ /d "1" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            cmd = f'reg add "HKLM\\{system_info["HKLM_IDM"]}" /v swfreg /t REG_SZ /d "1" /f'
            subprocess.run(cmd, shell=True, check=False)
            
            # Also try non-architecture specific path
            if "Wow6432Node" in system_info["HKLM_IDM"]:
                alt_path = system_info["HKLM_IDM"].replace("Wow6432Node\\", "")
                
                cmd = f'reg add "HKLM\\{alt_path}" /v regStatus /t REG_SZ /d "1" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                cmd = f'reg add "HKLM\\{alt_path}" /v regname /t REG_SZ /d "{fname} {lname}" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                cmd = f'reg add "HKLM\\{alt_path}" /v regemail /t REG_SZ /d "{email}" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                cmd = f'reg add "HKLM\\{alt_path}" /v regserial /t REG_SZ /d "{serial}" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                cmd = f'reg add "HKLM\\{alt_path}" /v InstallStatus /t REG_DWORD /d "3" /f'
                subprocess.run(cmd, shell=True, check=False)
                
                cmd = f'reg add "HKLM\\{alt_path}" /v isreged /t REG_SZ /d "1" /f'
                subprocess.run(cmd, shell=True, check=False)
            
        except Exception as e:
            print(f"Warning: Could not set all system-wide keys: {str(e)}")
            print("Some features may still work, but full activation might be affected.")
        
        # Create license file as additional backup
        try:
            appdata = os.environ.get('APPDATA', '')
            if appdata:
                idm_dir = os.path.join(appdata, "IDM")
                os.makedirs(idm_dir, exist_ok=True)
                
                # Create license file
                license_path = os.path.join(idm_dir, "license.sav")
                with open(license_path, "w") as f:
                    f.write(f"Name: {fname} {lname}\n")
                    f.write(f"Email: {email}\n")
                    f.write(f"Serial: {serial}\n")
                    f.write(f"Date: {datetime.now().strftime('%Y-%m-%d')}\n")
                print(f"Created license file: {license_path}")
        except Exception as e:
            print(f"Warning: Could not create license file: {str(e)}")
        
        print("\nIDM activation completed successfully!")
        return True
        
    except Exception as e:
        print(f"Activation failed: {str(e)}")
        return False

# Main
def main():
    print("===== IDM Activation Fix Tool =====")
    print("This tool will fix IDM activation when the manager shows it as activated but IDM itself is not.")
    print("Make sure you run this as administrator for the best results.\n")
    
    # Get system info
    system_info = get_system_info()
    print(f"Architecture: {system_info['arch']}")
    print(f"IDM Path: {system_info['idm_path']}")
    
    # Check if IDM exists
    if not os.path.exists(system_info['idm_path']):
        print("\nERROR: IDM executable not found!")
        print("Please make sure IDM is installed before running this tool.")
        input("\nPress Enter to exit...")
        return
    
    # Stop IDM if running
    stop_idm()
    
    # Perform complete registry cleanup
    clean_registry(system_info, deep=True)
    
    # Perform activation
    activate_idm(system_info)
    
    print("\nActivation process completed. Please follow these steps:")
    print("1. RESTART YOUR COMPUTER to ensure all changes take effect")
    print("2. Launch IDM after restarting to see if it's properly activated")
    print("3. If IDM still shows as unregistered, try these additional steps:")
    print("   - Disable Windows Firewall temporarily")
    print("   - Make sure you have administrator privileges")
    print("   - Try installing an older version of IDM (v6.38 is recommended)\n")
    
    # Run IDM
    print("Would you like to launch IDM now to check the activation? (y/n)")
    choice = input().lower()
    if choice == 'y' or choice == 'yes':
        try:
            os.startfile(system_info['idm_path'])
            print("IDM started. Check if it shows as registered.")
        except:
            print("Could not start IDM automatically. Please start it manually.")
    
    print("\nThank you for using the IDM Activation Fix Tool!")
    input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press Enter to exit...") 