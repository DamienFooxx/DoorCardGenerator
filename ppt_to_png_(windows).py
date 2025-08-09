import os
import comtypes.client

# WINDOWS SETUP, USING COMTYPES


def ppt_to_png(input_folder, output_folder):
    # Create PowerPoint application object
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    # Loop through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".ppt") or filename.endswith(".pptx"):
            input_path = os.path.join(input_folder, filename)

            # Check if the file exists
            if not os.path.exists(input_path):
                print(f"File not found: {input_path}")
                continue

            try:
                # Open the PowerPoint file
                presentation = powerpoint.Presentations.Open(input_path)

                # Create output folder if it doesn't exist
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)

                # Loop through each slide in the presentation
                for i, slide in enumerate(presentation.Slides, start=1):
                    # Set the output PNG file path
                    output_path = os.path.join(
                        output_folder, f"{os.path.splitext(filename)[0]}_slide{i}.png")

                    # Save the slide as PNG image
                    slide.Export(output_path, "PNG")

                # Close the presentation
                presentation.Close()

            except Exception as e:
                print(f"Error processing file: {input_path}")
                print(f"Error details: {e}")

    # Quit PowerPoint application
    powerpoint.Quit()


# Replace 'input_folder' with the path to the folder containing your PowerPoint files
input_folder = './24_25_doorcards_pptx'

# Replace 'output_folder' with the path to the folder where you want to save PNG images
output_folder = './24_25_doorcards_png'

ppt_to_png(input_folder, output_folder)
