import subprocess
import os

# MAC SETUP, USING SOFFICE AND BREW LIBREOFFICE


def ppt_to_png(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith((".ppt", ".pptx")):
            input_path = os.path.join(input_folder, filename)
            cmd = [
                "soffice",
                "--headless",
                "--convert-to", "png",
                "--outdir", output_folder,
                input_path
            ]
            subprocess.run(cmd, check=True)

# Change input and output folders accordingly


input_folder = './24_25_s2_pptx'
output_folder = './24_25_s2_doorcards_png'
ppt_to_png(input_folder, output_folder)
