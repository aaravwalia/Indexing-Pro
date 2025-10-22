import hashlib
import platform
import sys
from tkinter import messagebox # Using tkinter messagebox for consistency

# --- IMPORTANT: Secret Phrase for Activation Key Generation ---
# This MUST be IDENTICAL to the one in your Client Application (Indexing PRO).
# CHANGE THIS TO A STRONG, UNIQUE, AND SECRET PHRASE!
AUTHORIZED_SECRET_PHRASE = "MyCustomSecretPhraseForRenamerApp2025!"

def get_machine_fingerprint_for_key_gen():
    """
    Generates a unique fingerprint (System Code) for the machine (Windows-specific).
    This function is a copy of the one in the main app to ensure consistency.
    """
    if platform.system() != "Windows":
        messagebox.showerror("Platform Error", "This key generator is currently supported only on Windows.")
        return None

    try:
        import wmi 
        c = wmi.WMI()

        cpu_info = c.Win32_Processor()[0]
        cpu_id = cpu_info.ProcessorId if hasattr(cpu_info, 'ProcessorId') else ""

        board_info = c.Win32_BaseBoard()[0]
        board_serial = board_info.SerialNumber if hasattr(board_info, 'SerialNumber') else ""

        disk_serial = ""
        disk_info = c.Win32_DiskDrive()
        for disk in disk_info:
            if not disk.MediaType == "Removable Media" and disk.SerialNumber:
                disk_serial = disk.SerialNumber.strip()
                break
        
        # Fallback for disk serial if not found or if only removable media
        if not disk_serial and disk_info:
             disk_serial = disk_info[0].SerialNumber.strip() if hasattr(disk_info[0], 'SerialNumber') else ""

        mac_address = ""
        for nic in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
            if nic.MACAddress:
                mac_address = nic.MACAddress.replace(":", "").upper() # Clean MAC format
                break

        raw_fingerprint_string = f"{cpu_id}-{board_serial}-{disk_serial}-{mac_address}".strip()
        return hashlib.sha256(raw_fingerprint_string.encode('utf-8')).hexdigest().upper()

    except ImportError:
        messagebox.showerror("Dependency Error", "The 'wmi' module is required. Please install it using: pip install wmi")
        return None
    except Exception as e:
        messagebox.showerror("System Error", f"Could not generate machine ID: {e}\nEnsure you have necessary permissions.")
        return None

def generate_activation_key(system_code, secret_phrase):
    """
    Generates the activation key based on the system code and the secret phrase.
    This function MUST be identical to the one in your Client Application (Indexing PRO).
    """
    if not system_code:
        return None
    combined_string = f"{system_code}-{secret_phrase}"
    return hashlib.sha256(combined_string.encode('utf-8')).hexdigest().upper()

if __name__ == "__main__":
    print("--- Easy File Renamer Activation Key Generator ---")
    print("This utility can generate an activation key for a given System Code.")
    print("To generate a key for another user, ask them for their System Code (Machine ID).")
    
    # --- OPTION 1: Generate key for *your* machine (default behavior) ---
    # Uncomment the following lines if you want to generate a key for your own machine.
    # This is useful to test the activation process on your own system.
    # print("\n--- Generating Key for Your Own Machine ---")
    # system_code_for_self = get_machine_fingerprint_for_key_gen()
    # if system_code_for_self:
    #     print(f"Your System Code (Machine ID): {system_code_for_self}")
    #     activation_key_for_self = generate_activation_key(system_code_for_self, AUTHORIZED_SECRET_PHRASE)
    #     if activation_key_for_self:
    #         print(f"Generated Activation Key for your machine: {activation_key_for_self}")
    #     else:
    #         print("Could not generate activation key for your machine.")
    # else:
    #     print("\nFailed to retrieve your System Code.")

    # --- OPTION 2: Generate key for *another user* (use this for their machine code) ---
    print("\n--- Generating Key for Another User ---")
    user_provided_machine_code = input("Please paste the other user's System Code (Machine ID) here: ").strip().upper()

    if user_provided_machine_code:
        generated_key_for_user = generate_activation_key(user_provided_machine_code, AUTHORIZED_SECRET_PHRASE)
        
        if generated_key_for_user:
            print(f"\nSystem Code provided: {user_provided_machine_code}")
            print(f"Generated Activation Key for this System Code: {generated_key_for_user}")
            print("\nSend this generated key back to the user for them to activate their software.")
        else:
            print("\nCould not generate activation key for the provided System Code. Please check the code's format.")
    else:
        print("No System Code was entered. Cannot generate an activation key.")
    
    input("\nPress Enter to exit...") # Keep console open until user presses Enter