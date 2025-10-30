import win32com.client
import os
import numpy as np
import time
from params import Params
import matplotlib.pyplot as plt
import codev_helper as cvh
import os

# ==============================================================================
# Helper Functions
# ==============================================================================
    
def tilt2power(tilt):
    delta = -tilt*params.f0
    optical_power = delta*12*(params.eta - params.eta_air)/params.C0
    return optical_power


params = Params()


if __name__ == '__main__':
    # --- Configuration for CodeV session ---
    WORKING_DIR = os.getcwd() + "\\"
    LENS_FILE = WORKING_DIR + "system_with_camera" 
    RESULTS_DIR = WORKING_DIR + "sensitivity_analysis\\"

    # --- initialise variables ---
    distance = 0.5 # object distance in meters
    e_min = -2e-3
    e_max = 2e-3
    num_steps = 10
    epsilon = np.linspace(e_min, e_max, num_steps)

    # surfaces for each epsilon
    e1_surface = "S3"
    e2_surface = "S7"
    e3_surface = "S22"

    n_levels = 30
    cmap = 'viridis'

    # --- 2. Initialize and Start CODE V Session ---
    cv_session = None

    try: 
        # create the COM object to interact with CODE V
        cv_session = win32com.client.Dispatch("CodeV.Application")
        print("Successfully created CODE V session object.")

        # set the working directory and start the background process
        cv_session.StartingDirectory = WORKING_DIR
        cv_session.StartCodeV()
        print(f"CODE V background process started. Version: {cv_session.CodeVVersion}")

        # open the specified lens file
        print(f"Opening lens: {LENS_FILE}...")
        output = cv_session.Command(f"RES {LENS_FILE}")
        print(f"Lens opened. CODE V response: {output}")

        # Ensure the results directory exists
        if not os.path.exists(RESULTS_DIR):
            os.makedirs(RESULTS_DIR)
            print(f"Created results directory: {RESULTS_DIR}")

        # Initialize CodeVHelper
        cvHelper = cvh.CodeVHelper(cv_session, debug=True)


        # set the object distance 
        print(f"Setting object distance to {distance*1000} mm...")
        cv_session.Command(f"THI S0 {distance*1000}") 

        # get initial thickness of the surfaces
        s1_t = cvHelper.query_surf_thickness(e1_surface)
        s2_t = cvHelper.query_surf_thickness(e2_surface)
        s3_t = cvHelper.query_surf_thickness(e3_surface)
        print(f"Initial thicknesses - {e1_surface}: {s1_t} mm, {e2_surface}: {s2_t} mm, {e3_surface}: {s3_t} mm")

        # --- Main Processing Loop ---

        # --- e1 vs e2 sensitivity analysis ---
        E1, E2 = np.meshgrid(epsilon, epsilon)
        Pv_grid = np.zeros(E1.shape)

        for i in range(E1.shape[0]):
            for j in range(E1.shape[1]):
                e1 = E1[i, j]
                e2 = E2[i, j]

                # set the surface thicknesses
                cv_session.Command(f"THI {e1_surface} {s1_t + e1*1e3}")  # convert to mm
                cv_session.Command(f"THI {e2_surface} {s2_t + e2*1e3}")  # convert to mm

                # Apply vignetting
                vignetting_command = 'run "C:\\CODEV202203_SR1\\macro\\setvig.seq" 1e-07 0.1 100 NO YES ;GO'
                #print(f"  Applying vignetting: {vignetting_command}")
                cv_session.Command(vignetting_command)

                # perform automatic optimization
                optimization_command = "AUT; P YES; ERR CDV; MNC 5; DRA S1..30  NO; EFP ALL Y; EFT TA; GLA SO..I  NFK5 NSK16 NLAF2 SF4; GO"
                #print(f"  Performing optimization: {optimization_command}")
                cv_session.Command(optimization_command)

                # get the value of the tilt
                tilt = cvHelper.query_xypolynomial_coeff("S13", "C2")
                power = tilt2power(tilt)
                Pv_grid[i, j] = power

                # print progress bar
                progress = (i * E1.shape[1] + j + 1) / (E1.shape[0] * E1.shape[1]) * 100
                print(f"Sensitivity Analysis Progress: {progress:.2f} %", end='\r')
        
        # Plotting the sensitivity analysis result
        plt.figure(figsize=(8, 6))
        cp = plt.contourf(E1*1e3, E2*1e3, Pv_grid, levels=n_levels, cmap=cmap)
        plt.colorbar(cp, label='Optical Power (Diopters)')
        plt.xlabel('E1 (mm)')
        plt.ylabel('E2 (mm)')
        plt.title('Sensitivity Analysis: Optical Power vs E1 and E2')
        plt.savefig(os.path.join(RESULTS_DIR, 'sensitivity_analysis_e1_e2.png'))
        plt.show()

        min_val = np.min(Pv_grid)
        max_val = np.max(Pv_grid)

        print(f"\nSensitivity Analysis Complete. Optical Power Range: {min_val:.4f} D to {max_val:.4f} D")

        # --- e1 vs e3 sensitivity analysis ---
        # E1, E3 = np.meshgrid(epsilon, epsilon)
        # Pv_grid = np.zeros(E1.shape)

        # for i in range(E1.shape[0]):
        #     for j in range(E1.shape[1]):
        #         e1 = E1[i, j]
        #         e3 = E3[i, j]

        #         # set the surface thicknesses
        #         cv_session.Command(f"THI {e1_surface} {s1_t + e1*1e3}")  # convert to mm
        #         cv_session.Command(f"THI {e3_surface} {s3_t + e3*1e3}")  # convert to mm

        #         # Apply vignetting
        #         vignetting_command = 'run "C:\\CODEV202203_SR1\\macro\\setvig.seq" 1e-07 0.1 100 NO YES ;GO'
        #         #print(f"  Applying vignetting: {vignetting_command}")
        #         cv_session.Command(vignetting_command)

        #         # perform automatic optimization
        #         optimization_command = "AUT; P YES; ERR CDV; MNC 5; DRA S1..30  NO; EFP ALL Y; EFT TA; GLA SO..I  NFK5 NSK16 NLAF2 SF4; GO"
        #         #print(f"  Performing optimization: {optimization_command}")
        #         cv_session.Command(optimization_command)

        #         # get the value of the tilt
        #         tilt = query_xrplynomial_coeff(cv_session, "S13", "C2")
        #         power = tilt2power(tilt)
        #         Pv_grid[i, j] = power

        #         # print progress bar
        #         progress = (i * E1.shape[1] + j + 1) / (E1.shape[0] * E1.shape[1]) * 100
        #         print(f"Sensitivity Analysis Progress: {progress:.2f} %", end='\r')

        # # Plotting the sensitivity analysis result
        # plt.figure(figsize=(8, 6))
        # cp = plt.contourf(E1*1e3, E3*1e3, Pv_grid, levels=n_levels, cmap=cmap)
        # plt.colorbar(cp, label='Optical Power (Diopters)')
        # plt.xlabel('E1 (mm)')
        # plt.ylabel('E3 (mm)')
        # plt.title('Sensitivity Analysis: Optical Power vs E1 and E3')
        # plt.savefig(os.path.join(RESULTS_DIR, 'sensitivity_analysis_e1_e3.png'))
        # plt.show()

        # --- e2 vs e3 sensitivity analysis ---
        # E2, E3 = np.meshgrid(epsilon, epsilon)
        # Pv_grid = np.zeros(E2.shape)

        # for i in range(E2.shape[0]):
        #     for j in range(E2.shape[1]):
        #         e2 = E2[i, j]
        #         e3 = E3[i, j]

        #         # set the surface thicknesses
        #         cv_session.Command(f"THI {e2_surface} {s2_t + e2*1e3}")  # convert to mm
        #         cv_session.Command(f"THI {e3_surface} {s3_t + e3*1e3}")  # convert to mm

        #         # Apply vignetting
        #         vignetting_command = 'run "C:\\CODEV202203_SR1\\macro\\setvig.seq" 1e-07 0.1 100 NO YES ;GO'
        #         #print(f"  Applying vignetting: {vignetting_command}")
        #         cv_session.Command(vignetting_command)

        #         # perform automatic optimization
        #         optimization_command = "AUT; P YES; ERR CDV; MNC 5; DRA S1..30  NO; EFP ALL Y; EFT TA; GLA SO..I  NFK5 NSK16 NLAF2 SF4; GO"
        #         #print(f"  Performing optimization: {optimization_command}")
        #         cv_session.Command(optimization_command)
        #         # get the value of the tilt
        #         tilt = query_xrplynomial_coeff(cv_session, "S13", "C2")
        #         power = tilt2power(tilt)
        #         Pv_grid[i, j] = power
        #         # print progress bar
        #         progress = (i * E2.shape[1] + j + 1) / (E2.shape[0] * E2.shape[1]) * 100
        #         print(f"Sensitivity Analysis Progress: {progress:.2f} %", end='\r')
        # # Plotting the sensitivity analysis result
        # plt.figure(figsize=(8, 6))
        # cp = plt.contourf(E2*1e3, E3*1e3, Pv_grid, levels=n_levels, cmap=cmap)
        # plt.colorbar(cp, label='Optical Power (Diopters)')
        # plt.xlabel('E2 (mm)')
        # plt.ylabel('E3 (mm)')
        # plt.title('Sensitivity Analysis: Optical Power vs E2 and E3')
        # plt.savefig(os.path.join(RESULTS_DIR, 'sensitivity_analysis_e2_e3.png'))
        # plt.show()

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # --- Crlan Up and Close session ---
        if cv_session:
            cv_session.StopCodeV()
            print("\nCODE V session stopped.")
            cv_session = None
