import win32com.client
import os
import numpy as np
import time
from params import Params

# ==============================================================================
# Helper Functions
# ==============================================================================

params = Params()

def format_for_filename(value: float) -> str:
    """
    Formats a float into the specified '{integer}p{decimal}' string format.
    Example: 12.34 -> "12p34"
    """
    if value < 0:
        sign = "-" # Using 'm' for minus
        value = abs(value)
    else:
        sign = ""
    
    integer_part = int(value)
    # Get the decimal part as a string, remove the "0." part and take the first two digits
    decimal_part = f"{(value - integer_part):.2f}"[2:]
    return f"{sign}{integer_part}p{decimal_part}"

def calculate_tilt(optical_power: float) -> float:
    """
    Placeholder function to calculate the tilt value from optical power.
    
    !!! IMPORTANT !!!
    You will need to replace the formula in this function with your actual calculation.
    """
    # --- Replace this placeholder formula ---
    delta = params.C0/(12*(params.eta - params.eta_air))*optical_power
    tilt_value = -delta/params.f0
    # ------------------------------------
    
    return tilt_value

# ==============================================================================
# Main Script Logic
# ==============================================================================

if __name__ == '__main__':
    # --- 1. Define Input Data and Configuration ---
    
    # Array of tasks to perform. Each element is a dictionary.
    tasks = [
        {'d': 0.5, 'p_t': 1.32},
        # Add more tasks as needed
    ]

    # Configuration for the CODE V session
    WORKING_DIR = "C:\\users\\crendon\\documents\\github\\codev_python_com\\"
    LENS_FILE = WORKING_DIR + "system_with_camera"  # IMPORTANT: Change to your lens file
    RESULTS_DIR = WORKING_DIR + "usaf_slm_calibration\\"

    # --- 2. Initialize and Start CODE V Session ---
    cv_session = None
    try:
        # Create the COM object to interact with CODE V
        cv_session = win32com.client.Dispatch("CodeV.Application")
        print("Successfully created CODE V session object.")

        # Set working directory and start the background process
        cv_session.StartingDirectory = WORKING_DIR
        cv_session.StartCodeV()
        print(f"CODE V background process started. Version: {cv_session.CodeVVersion}")

        # Open the specified lens file
        print(f"Opening lens: {LENS_FILE}...")
        cv_session.Command(f"RES {LENS_FILE}")

        # Ensure the results directory exists
        if not os.path.exists(RESULTS_DIR):
            os.makedirs(RESULTS_DIR)
            print(f"Created results directory: {RESULTS_DIR}")

        # --- 3. Main Processing Loop ---
        for task in tasks:
            distance = task['d']
            theoretical_power = task['p_t']
            
            print(f"\n--- Starting task for distance: {distance} ---")

            # place the object at the specified distance
            print(f"Setting object distance to {distance*1000} mm...")
            cv_session.Command(f"THI S0 {distance*1000}") 

            # Create the array of optical powers to test
            start_power = theoretical_power - 0.5
            end_power = theoretical_power + 0.5
            step = 0.05
            powers = np.arange(start_power, end_power + step, step)
            for current_power in powers:
                # Calculate the tilt value using the placeholder function
                tilt = calculate_tilt(current_power)
                
                # Format values for the filename
                dist_str = format_for_filename(distance)
                current_power_str = format_for_filename(current_power)
                
                # A. Apply the tilt value to the X coefficient of surface S13
                # Note: This assumes 'X' is a valid coefficient alias for your surface type.
                sco_command = f"SCO S13 X {tilt:.6f}"
                print(f"  Setting tilt: {sco_command}")
                cv_session.Command(sco_command)

                # B. Construct and run the IMS command block
                output_file = os.path.join(RESULTS_DIR, f"d_{dist_str}_slm_{current_power_str}")
                
                ims_command = f"""
                IMS;
                TGR 1024;
                OBJ CV_IMAGE:USAF1951_8MP.bmp;
                PMX 15;
                PMY 15;
                DEX 3.75e-3;
                DEY 3.75e-3;
                SVI BMP "{output_file}";
                GO;
                """
                
                print(f"  Running IMS, saving to {output_file}.bmp...")
                cv_session.Command(ims_command)
            

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # --- 4. Clean Up and Close Session ---
        if cv_session:
            cv_session.StopCodeV()
            print("\nCODE V session stopped.")
            cv_session = None