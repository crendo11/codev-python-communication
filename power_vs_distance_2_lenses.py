"""
this script will create the simulation of having two lenses and vary the distance between them
and plot the optical power vs distance for the combined error
"""

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
    epsilon = 2

    # surfaces for each epsilon
    e1_surface = "S3"           # Lens 1 to Lens 2
    e6_surface = "S22"          # Lens 4 to lens 5
    e8_surface = "S29"          # Lens 6 to camera's sensor

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
        s6_t = cvHelper.query_surf_thickness(e6_surface)
        s8_t = cvHelper.query_surf_thickness(e8_surface)

        surfaces_thickness = [s1_t, s6_t, s8_t]

        print(f"Initial thicknesses - {e1_surface}: {s1_t} mm, {e6_surface}: {s6_t} mm, {e8_surface}: {s8_t} mm")


        # --- Main Processing Loop ---
        # reset thicknesses
        cvHelper.set_surf_thickness(e1_surface, s1_t)
        cvHelper.set_surf_thickness(e6_surface, s6_t)
        cvHelper.set_surf_thickness(e8_surface, s8_t)
        

        # set distance between lenses
        cvHelper.set_surf_thickness(e1_surface, s1_t + 1)  # convert to mm
        cvHelper.set_surf_thickness(e8_surface, s8_t + 1)  # convert to mm
        cvHelper.set_surf_thickness(e6_surface, s6_t + 3)  # convert to mm
        i = 0
        powers = []
        plt.figure()
        for d in distances:
            
            # set the object distance
            #print(f"Setting object distance to {d*1000} mm...")
            cv_session.Command(f"THI S0 {d*1000}")

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

            # print percentage completed 
            progress = (distances.index(d) + 1) / len(distances) * 100
            # print(f"Distance Analysis Progress: {progress:.2f} %", end='\r')
            print(power)

        plt.scatter(distances, powers, label=f"Epsilon: {epsilon} um")

        # plot the theoretical curve
        d = np.linspace(0.4, 4, 100)  # distance in meters
        P = 9 / (16 * (d - 0.075))
        plt.plot(d, P, 'r--', label="Theoretical Curve")
        plt.xlabel("Distance (m)")
        plt.ylabel("Optical Power (D)")
        plt.title("Optical Power vs Distance for Combined S3 and S22")
        plt.grid()
        plt.legend()
        i += 1
        plt.show()



                



    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # --- Crlan Up and Close session ---
        if cv_session:
            cv_session.StopCodeV()
            print("\nCODE V session stopped.")
            cv_session = None
