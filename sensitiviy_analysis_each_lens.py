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
    distances = [0.4, 0.5, 0.6, 0.7, 0.8, 2, 3.75]
    e_min = -3e-3
    e_max = 3e-3
    num_steps = 15
    epsilon = np.linspace(e_min, e_max, num_steps)

    # surfaces for each epsilon
    e1_surface = "S3"           # Lens 1 to Lens 2
    e2_surface = "S7"           # Lens 2 to Lohmann lens
    e3_surface = "S9"           # Lohmann to lens 3
    e4_surface = "S12"          # Lens 3 to SLM
    e5_surface = "S13"          # Lens 3 to SLM e5 must be  e5 = -e4 S13 now has a pickup to S12
    e6_surface = "S22"          # Lens 4 to lens 5
    e7_surface = "S26"          # Lens 5 to lens 6
    e8_surface = "S29"          # Lens 6 to camera's sensor

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
        cvHelper = cvh.CodeVHelper(cv_session)

        # get initial thickness of the surfaces
        s1_t = cvHelper.query_surf_thickness(e1_surface)
        s2_t = cvHelper.query_surf_thickness(e2_surface)
        s3_t = cvHelper.query_surf_thickness(e3_surface)
        s4_t = cvHelper.query_surf_thickness(e4_surface)
        s5_t = cvHelper.query_surf_thickness(e5_surface)
        s6_t = cvHelper.query_surf_thickness(e6_surface)
        s7_t = cvHelper.query_surf_thickness(e7_surface)
        s8_t = cvHelper.query_surf_thickness(e8_surface)

        surfaces_thickness = [s1_t, s2_t, s3_t, s4_t, s6_t, s7_t, s8_t]

        print(f"Initial thicknesses - {e1_surface}: {s1_t} mm, {e2_surface}: {s2_t} mm, {e3_surface}: {s3_t} mm, {e4_surface}: {s4_t} mm, {e5_surface}: {s5_t} mm, {e6_surface}: {s6_t} mm, {e7_surface}: {s7_t} mm, {e8_surface}: {s8_t} mm")

        # print lens 
        cvHelper.plot_lens("initial_lens")
        epsilon = np.linspace(e_min, e_max, num_steps)

        # --- Main Processing Loop ---
        i = 0
        for e_surface in [e1_surface, e2_surface, e3_surface, e4_surface, e6_surface, e7_surface, e8_surface]:
            
            # reset thicknesses
            cvHelper.set_surf_thickness(e1_surface, s1_t)
            cvHelper.set_surf_thickness(e2_surface, s2_t)
            cvHelper.set_surf_thickness(e3_surface, s3_t)
            cvHelper.set_surf_thickness(e4_surface, s4_t)
            cvHelper.set_surf_thickness(e5_surface, s5_t)
            cvHelper.set_surf_thickness(e6_surface, s6_t)
            cvHelper.set_surf_thickness(e7_surface, s7_t)
            cvHelper.set_surf_thickness(e8_surface, s8_t)
            
            plt.figure()
            
            for dist in distances:
                # set the object distance 
                print(f"Setting object distance to {dist*1000} mm...")
                cvHelper.set_surf_thickness("S0", dist*1000)  # convert to mm

                powers = [] 
                for e in epsilon:
                    # set the thickness for surface 
                    cvHelper.set_surf_thickness(e_surface, surfaces_thickness[i] + e*1e3)  # convert to mm

                    if i == 3:  # e4_surface
                        # set e5 to be -e4
                        cvHelper.set_surf_thickness(e5_surface, s5_t - e*1e3)  # convert to mm

                    # apply vignetting
                    cvHelper.apply_vignetting()

                    # perform automatic optimization
                    optimization_command = "AUT; P YES; ERR CDV; MNC 5; DRA S1..30  NO; EFP ALL Y; EFT TA; GLA SO..I  NFK5 NSK16 NLAF2 SF4; GO"
                    #print(f"  Performing optimization: {optimization_command}")
                    cv_session.Command(optimization_command)

                    # get the value of the tilt
                    tilt = cvHelper.query_xypolynomial_coeff("S13", "C2")
                    power = tilt2power(tilt)

                    # store
                    powers.append(power)

                # save matrices
                np.savez(os.path.join(RESULTS_DIR, f"sensitivity_{e_surface}_dist_{int(dist*1000)}mm.npz"), epsilon=epsilon, powers=powers, dist=dist)


                plt.plot(epsilon*1e3, powers, marker='o', label=r'$d_O' + f'= {dist*1000} mm')
                plt.title(r"Surface" + f"{e_surface}")
                plt.xlabel('Epsilon (mm)')
                plt.ylabel('Optical Power (Diopters)')
                plt.grid(True)
            
            plt.legend()
            plot_filename = os.path.join(RESULTS_DIR, f"sensitivity_{e_surface}.png")
            plt.savefig(plot_filename)
            i += 1

            # print percentage complete
            progress = (i) / 7 * 100
            print(f"Sensitivity Analysis Progress: {progress:.2f} %")


                



    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # --- Crlan Up and Close session ---
        if cv_session:
            cv_session.StopCodeV()
            print("\nCODE V session stopped.")
            cv_session = None
