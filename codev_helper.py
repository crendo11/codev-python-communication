"""
This library contains helper functions for interacting with code V.
The functions here will receive the code V session object as a parameter
and perform specific tasks.
"""


class CodeVHelper:

    # init
    def __init__(self, cv_session, debug=False):
        self.cv_session = cv_session
        self.debug = debug

    def plot_lens(self, plot_filename):
        # Set the graphics output to a file
        self.cv_session.Command(f"GRA {plot_filename}")
        # Generate the 2D plot
        self.cv_session.Command("VIE; PLC; GO")
        print(f"Plot saved to {plot_filename}.plt")

        # convert the .plt file to .jpg
        self.cv_session.Command(f"GCV JPG {plot_filename}.plt")
        print(f"Converted {plot_filename}.plt to {plot_filename}.jpg")

    def query_surf_thickness(self, surface):
        command = f"?THI {surface}"
        if self.debug:
            print(f"Executing command: {command}")
        output = self.cv_session.Command(command)
        if output:
            value = float(output.split("=")[1].split("\r")[0])
            if self.debug:
                print(f"Output: {output}")
            return value
        else:
            return None
    
    def query_xypolynomial_coeff(self, surface, order):
        command = f"?SCO {surface} {order}"
        if self.debug:
            print(f"Executing command: {command}")
        output = self.cv_session.Command(command)
        if output:
            if self.debug:
                print(f"Output: {output}")
            value = float(output.split("=")[1].split("\r")[0])
            return value
        else:
            return None
        
    def set_surf_thickness(self, surface, new_thickness):
        command = f"THI {surface} {new_thickness}"
        if self.debug:
            print(f"Executing command: {command}")
        output = self.cv_session.Command(command)
        if self.debug:
            print(f"Output: {output}")

    def apply_vignetting(self):
        vignetting_command = 'run "C:\\CODEV202203_SR1\\macro\\setvig.seq" 1e-07 0.1 100 NO YES ;GO'
        if self.debug:
            print(f"  Applying vignetting: {vignetting_command}")
        output = self.cv_session.Command(vignetting_command)
        if self.debug:
            print(f"Output: {output}")
        return output