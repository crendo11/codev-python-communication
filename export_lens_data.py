import win32com.client
import csv
import os

# ================= CONFIGURATION =================
# REPLACE THIS PATH with the actual path to your .len file.
# Use 'r' before the string to handle backslashes correctly.
LENS_FILE_PATH = r"./system_with_camera.len" 

# The name of the output CSV file to create
OUTPUT_CSV = "lens_data_export.csv"
# =================================================

def export_lens_data():
    cv = None
    try:
        # 1. Initialize the CODE V Connection
        print("Connecting to CODE V...")
        # [cite_start]API Reference: Instantiating the Client Object [cite: 305-312]
        cv = win32com.client.Dispatch("CodeV.Application")
        
        # Start the application (or hook into existing one)
        # [cite_start]API Reference: Method StartCodeV [cite: 887-892]
        cv.StartCodeV()

        # 2. Load the Lens File
        print(f"Loading lens: {LENS_FILE_PATH}...")
        # [cite_start]API Reference: Method Command [cite: 1022-1025]
        # We use the standard RES command to load the file
        cv.Command(f'RES "{LENS_FILE_PATH}"')

        # 3. Get the Total Number of Surfaces
        # We query the (NUM S) database item.
        # [cite_start]API Reference: Method EvaluateExpression [cite: 1034-1045]
        num_surfaces_str = cv.EvaluateExpression("(NUM S)")
        num_surfaces = int(float(num_surfaces_str))
        
        print(f"Found {num_surfaces} surfaces (Object to Image). Exporting...")

        # 4. Extract Data and Write to CSV
        with open(OUTPUT_CSV, mode='w', newline='') as file:
            writer = csv.writer(file)
            
            # Write the Header Row
            writer.writerow(["Surface", "Radius", "Thickness", "Glass", "Index"])
            
            # Loop from 0 (Object) to num_surfaces (Image)
            # We use range(num_surfaces + 1) because Python ranges are exclusive at the top
            for i in range(num_surfaces + 1):
                
                # Query LDM Database Items for the current surface
                # [cite_start]Prompting Guide: Surface Data Items (RDY, THI, GLA, IND) [cite: 349-353, 4084]
                rdy = cv.EvaluateExpression(f"(RDY S{i})")
                thi = cv.EvaluateExpression(f"(THI S{i})")
                gla = cv.EvaluateExpression(f"(GLA S{i})")
                ind = cv.EvaluateExpression(f"(IND S{i})")
                
                # Write the row to the CSV file
                writer.writerow([i, rdy, thi, gla, ind])
                
        print(f"Success! Data exported to: {os.path.abspath(OUTPUT_CSV)}")

    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
        print("Make sure CODE V is installed and the lens file path is correct.")

    finally:
        # 5. Clean Up
        # It is good practice to release the COM object. 
        # Uncomment cv.StopCodeV() if you want the script to close CODE V automatically.
        if cv:
            # cv.StopCodeV() 
            pass

if __name__ == "__main__":
    export_lens_data()