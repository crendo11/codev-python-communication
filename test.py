
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

params = Params()
    

def tilt2power(tilt):
    delta = -tilt*params.f0
    optical_power = delta*12*(params.eta - params.eta_air)/params.C0
    return optical_power

def rotate_lohmann_lens(cv_session, surface, theta = 0, c0 = 1/0.00013312, debug=False):

    # compute coefficients
    x3 = 1/c0 * (np.sin(theta)**3 + np.cos(theta)**3)
    x2y = 3/c0 * (np.sin(theta)**2 * np.cos(theta) - np.cos(theta)**2 * np.sin(theta))
    xy2 = 3/c0 * (np.sin(theta)**2 * np.cos(theta) + np.sin(theta) * np.cos(theta)**2)
    y3 = 1/c0 * (np.cos(theta)**3 - np.sin(theta)**3)

    lohmann_surf = surface
    x3_set_command = f"SCO {lohmann_surf} C7 {str(x3)}"
    x2y_set_command = f"SCO {lohmann_surf} C8 {str(x2y)}"
    xy2_set_command = f"SCO {lohmann_surf} C9 {str(xy2)}"
    y3_set_command = f"SCO {lohmann_surf} C10 {str(y3)}"

    # set coefficients
    output = cv_session.Command(x3_set_command)
    if debug:
        print(f"Setting {x3_set_command}, output: {output}")
    output = cv_session.Command(x2y_set_command)
    if debug:
        print(f"Setting {x2y_set_command}, output: {output}")
    output = cv_session.Command(xy2_set_command)
    if debug:
        print(f"Setting {xy2_set_command}, output: {output}")
    output = cv_session.Command(y3_set_command)
    if debug:
        print(f"Setting {y3_set_command}, output: {output}")

def rotate_SLM(cv_session, dummy_surface, theta = 0, debug=False):

    a_coeff = "C2"
    x2_coeff = "C4"
    x3_coeff = "C7"

    # calculate the constants from the angle
    sin_th = np.sin(theta)
    cos_th = np.cos(theta)

    # set x2 = a * sin(th) + cos(th)
    command_x2 = f"PIK SCO {x2_coeff} {dummy_surface} SCO {a_coeff} {dummy_surface} {str(sin_th)} {str(cos_th)}"
    output = cv_session.Command(command_x2)
    if debug:
        print(f"Setting {command_x2}, output: {output}")

    # set x3 = a * cos(th) - sin(th)
    command_x3 = f"PIK SCO {x3_coeff} {dummy_surface} SCO {a_coeff} {dummy_surface} {str(cos_th)} {str(-sin_th)}"
    output = cv_session.Command(command_x3)
    if debug:
        print(f"Setting {command_x3}, output: {output}")

    

def get_power_vs_distance_with_epsilon(cv_session, surface1, surface2, surface3, st1, st2, st3, theta, distances):
    
    cvHelper.set_surf_thickness(surface1, st1)  
    cvHelper.set_surf_thickness(surface2, st2)  
    cvHelper.set_surf_thickness(surface3, st3)

    # roate the lohmann lenses
    lohmann_surf = "S9"
    rotate_lohmann_lens(cv_session, lohmann_surf, theta)
    
    powers = []
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
        print(power)

    return powers


def get_power_vs_distance_with_SLM_tilt(cv_session, surface1, surface2, surface3, st1, st2, st3, theta_lohmann, theta_SLM, distances):
    
    cvHelper.set_surf_thickness(surface1, st1)  
    cvHelper.set_surf_thickness(surface2, st2)  
    cvHelper.set_surf_thickness(surface3, st3)

    # roate the lohmann lenses
    lohmann_surf = "S9"
    rotate_lohmann_lens(cv_session, lohmann_surf, theta_lohmann)
    # rotate the SLM
    dummy_surface = "S14"
    rotate_SLM(cv_session, dummy_surface, theta_SLM)
    
    powers = []
    for d in distances:
        
        # set the object distance
        #print(f"Setting object distance to {d*1000} mm...")
        cv_session.Command(f"THI S0 {d*1000}")

        # apply vignetting
        cvHelper.apply_vignetting()

        # perform automatic optimization
        optimization_command = get_optimization_with_SLM_tilt_command()
        print(f"  Performing optimization: {optimization_command}")
        output = cv_session.Command(optimization_command)
        print(output)

        # get the value of the tilt
        tilt = cvHelper.query_xypolynomial_coeff("S14", "C2")
        power = tilt2power(tilt)

        # store
        powers.append(power)

        # print percentage completed 
        print(power)

    return powers

def get_optimization_with_SLM_tilt_command():
    return "AUT; @x_coeff == 1/(SCO S14 C4) * (SCO S14 C7) - (SCO S13 C2); @y_coeff == (SCO S14 C2) / (SCO S14 C4) - (SCO S13 C3); @x_coeff = 0; @y_coeff = 0; SCO S14 C2 < 1e-5; STP YES; ERR CDV; MNC 5; DRA S1..30  NO; EFP ALL Y; EFT TA; GLA SO..I  NFK5 NSK16 NLAF2 SF4; GO"

# --- Configuration for CodeV session ---
WORKING_DIR = os.getcwd() + "\\"
LENS_FILE = WORKING_DIR + "system_with_camera_tilt_SLM" 
RESULTS_DIR = WORKING_DIR + "sensitivity_analysis\\"

# --- initialise variables ---
distances = [0.5, 0.6, 0.7, 0.8, 2, 3.75]
epsilon = 2

c0 = 1/0.00013312  # in mm

# surfaces for each epsilon
e1_surface = "S3"           # Lens 1 to Lens 2
e2_surface = "S7"           # Lens 2 to Lohmann lens
e3_surface = "S9"           # Lohmann to lens 3
e4_surface = "S12"          # Lens 3 to SLM
e5_surface = "S13"          # Lens 3 to SLM e5 must be  e5 = -e4
e6_surface = "S22"          # Lens 4 to lens 5
e7_surface = "S26"          # Lens 5 to lens 6
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
    cvHelper = cvh.CodeVHelper(cv_session, debug=True)

    # get initial thickness of the surfaces
    s1_t = cvHelper.query_surf_thickness(e1_surface)
    s2_t = cvHelper.query_surf_thickness(e2_surface)
    s3_t = cvHelper.query_surf_thickness(e3_surface)
    s4_t = cvHelper.query_surf_thickness(e4_surface)
    s5_t = cvHelper.query_surf_thickness(e5_surface)
    s6_t = cvHelper.query_surf_thickness(e6_surface)
    s7_t = cvHelper.query_surf_thickness(e7_surface)
    s8_t = cvHelper.query_surf_thickness(e8_surface)

    surfaces_thickness = [s1_t, s6_t, s8_t]

    print(f"Initial thicknesses - {e1_surface}: {s1_t} mm, {e6_surface}: {s6_t} mm, {e8_surface}: {s8_t} mm")
    
    # reset thicknesses
    cvHelper.set_surf_thickness(e1_surface, s1_t)
    cvHelper.set_surf_thickness(e6_surface, s6_t)
    cvHelper.set_surf_thickness(e8_surface, s8_t)
    

    # --- Main Processing Loop ---

    #  experimental result
    p_SLM_m = [1.02, 0.67, 0.5, 0.25, -0.6, -1]

    # plot the theoretical curve
    d = np.linspace(0.4, 4, 100)  # distance in meters
    P = 9 / (16 * (d - 0.075))


    plt.figure(figsize=(10, 6))
    
    # for i in np.arange(-20, 20, 2):
    for i in [0]:
        rot = i
        powers = get_power_vs_distance_with_SLM_tilt(cv_session, e1_surface, e6_surface, e8_surface, s1_t + 0, s6_t + 0, s8_t + 0, 0, np.deg2rad(rot), distances)
        plt.plot(distances, powers, 'o--', label=f"rotated: {rot} deg")


    
    plt.scatter(distances, p_SLM_m, marker='o', color='g', label="Experimental SLM Data")
    plt.plot(d, P, 'r--', label="Theoretical Curve")
    plt.xlabel("Distance (m)")
    plt.ylabel("Optical Power (D)")
    plt.title("Optical Power vs Distance for Rotated Lohmann Lens")
    plt.grid()
    plt.legend()

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # --- Crlan Up and Close session ---
    if cv_session:
        cv_session.StopCodeV()
        print("\nCODE V session stopped.")
        cv_session = None