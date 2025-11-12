MultiCap Dx Web Diagnostics Interface

This repository contains the supplementary Python source code for the MultiCap Dx platform. This is a custom Flask-based web application designed to manage serial communication with the electronically integrated sensor, acquire 160x160 pixel capacitance data, and provide real-time Region of Interest (ROI) analysis for multi-viral antigen diagnosis.

This code supports the findings presented in the paper, "Bubble-Induced 3D Capacitance Profiling for Quantitative, Multi-Viral Antigen Detection at the Point of Care."

1. Prerequisites

To run this application, you must have the following:

Python 3.x (version 3.8 or higher is recommended).

The required Python libraries (listed below).

A connected MCU device for serial data acquisition.

2. Installation

Clone this repository to your local machine:

# Use the actual URL provided by GitHub after repository creation
git clone [Your GitHub Repository URL Here]
cd [repository-name]


Install all necessary dependencies using the provided requirements.txt file:

pip install -r requirements.txt


Required Libraries

Library

Role

Flask

Web server framework setup

numpy

Handling image and data arrays

Pillow

Image (PNG) transformation and processing

pyserial

Serial communication with the device (MCU)

xlsxwriter

For potential future Excel saving capability

3. Configuration (Crucial Step!)

The application needs to know which communication port the device is connected to.

Open the app.py file in a text editor.

Locate the Configuration section near the top of the file and modify the PORT_NAME variable to match the actual COM port (e.g., COM3 on Windows, or /dev/ttyUSB0 on Linux/macOS) of your connected MCU device.

# app.py snippet:
# ---------------------- Configuration ----------------------
PORT_NAME    = "COM7"  # <--- MUST be changed to the device's actual COM port
BAUDRATE     = 115200
# ...


4. Usage

Start the Flask server from your terminal:

python app.py


You should see a message indicating the server is running.

Open your web browser and navigate to:
http://127.0.0.1:5050

Use the UI to perform diagnostics:

Click the RUN button to initiate serial communication, capture the 160x160 pixel capacitance image, and display it on the UI.

(Optional) Use the ROI adjustment buttons (↑ ↓ ← →) to fine-tune the center coordinates of the four analysis regions.

Click the Extract & Analyze button to perform quantitative analysis based on the predefined cutoffs.

5. Citation

If you use this code in your own research, please cite the original paper:

"Bubble-Induced 3D Capacitance Profiling for Quantitative, Multi-Viral Antigen Detection at the Point of Care."
