import win32com.client
import os

# --- 1. Connect to the CODE V Application ---
# Create an instance of the CODE V application object using its ProgID.
# This is the standard way to start a COM session from Python.
try:
    cv_session = win32com.client.Dispatch("CodeV.Application")
    print("Successfully created CODE V session object.")
except Exception as e:
    print(f"Failed to create COM object. Make sure CODE V is installed. Error: {e}")
    exit()


# define a function to plot the lens
def plot_lens(cv_session, plot_filename):
    # Set the graphics output to a file
    cv_session.Command(f"GRA {plot_filename}")
    # Generate the 2D plot
    cv_session.Command("VIE; PLC; GO")
    print(f"Plot saved to {plot_filename}.plt")

    # convert the .plt file to .jpg
    cv_session.Command(f"GCV JPG {plot_filename}.plt")
    print(f"Converted {plot_filename}.plt to {plot_filename}.jpg")


# --- 2. Start the CODE V Background Process ---
# Set the working directory and start the session.
cv_session.StartingDirectory = "C:\\users\\crendon\\documents\\github\\codev_python_com"
cv_session.StartCodeV()
print(f"CODE V background process started.")
print(f"Working Directory: {cv_session.StartingDirectory}")

# --- 3. Send Commands to CODE V ---
# We will send a series of commands to perform the requested actions.

# A. Restore a sample lens file (e.g., the double gauss lens)
lens_file = "./system_with_camera.len"
print(f"Opening lens: {lens_file}...")
result_open = cv_session.Command(f"RES {lens_file}")

# B. Prepare to save the plot and generate it
# The 'GRA' command directs the next graphical output to a file.
# The 'VIE; PLC; GO' sequence then generates the 2D plot.
plot_filename_no_ext = "Testing_plot"

plot_lens(cv_session, plot_filename_no_ext)


# move the object to infinity
print("Moving object to infinity...")
result_move = cv_session.Command("THI S0 1E14")

plot_name_obj_infinity = "Testing_plot_object_infinity"
plot_lens(cv_session, plot_name_obj_infinity)



# --- 4. Close the CODE V Session ---
# It's crucial to stop the session to release the license and process.
cv_session.StopCodeV()
print("CODE V session stopped.")

# Clean up the Python COM object
cv_session = None
print("Script finished.")



