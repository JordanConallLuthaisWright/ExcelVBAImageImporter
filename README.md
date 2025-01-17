
# Excel Image Importer with VBA

## Overview
The **Excel Image Importer** is a VBA macro designed to automate the process of importing images into Excel. This tool allows users to organize images in a structured manner, placing each image in its own cell with a comment section in the adjacent column. The macro ensures professional formatting with consistent sizing, aspect ratio preservation, and borders for better organization and presentation.

## Features
- **Batch Import:** Automatically imports all images from a specified folder.
- **Custom Sizing:** Resizes images to fit cells while maintaining their original aspect ratio.
- **Professional Borders:** Adds clean borders around the image and comment cells.
- **Comment Section:** Creates a dedicated column for user comments next to each image.
- **Multi-Format Support:** Supports popular image formats like `.jpg`, `.jpeg`, `.png`, `.bmp`, and `.gif`.

## Use Cases
This tool can be used for a variety of purposes:
- Documenting and annotating **building defects** (e.g., for architects or contractors).
- Managing **product images** for inventory or presentation purposes.
- Creating structured **visual reports** for business or personal projects.

---

## Installation Instructions

### Step 1: Download the VBA Code
1. Go to the [GitHub repository](#) (insert your GitHub link here).
2. Download the `ImportImages.bas` file by clicking on it and selecting **Download**.

### Step 2: Import the VBA Code into Excel
1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, go to `File > Import File...`.
4. Select the downloaded `ImportImages.bas` file and click **Open**.
5. Close the VBA editor by pressing `Alt + Q`.

### Step 3: Run the Macro
1. In Excel, press `Alt + F8` to open the **Macro dialog box**.
2. Select `ImportImagesWithResizingAndBordersFixed` (or the macro name) from the list.
3. Click **Run** and follow the prompts to specify the folder containing your images.

---

## How It Works
1. The macro scans a specified folder for images (supports `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`).
2. Each image is inserted into **Column A**, resized to a manageable size (100x100 pixels by default).
3. A comment cell is added in **Column B** next to the image for annotations.
4. Borders are applied to both the image and comment cells for a clean, professional look.

---

## Example Output
| **Image (Column A)** | **Comments (Column B)** |
|-----------------------|-------------------------|
| ![Image](#)           | Enter comment here      |
| ![Image](#)           | Enter comment here      |

---

## Customization
You can customize the following aspects of the macro:
- **Image Size:** Modify the `PicMaxWidth` and `PicMaxHeight` constants in the VBA code to adjust the image dimensions.
- **Folder Path:** Set a default folder path or allow dynamic folder selection during runtime.
- **Borders:** Change the border style or color by editing the `.Borders` section of the code.

---

## Troubleshooting
### Common Issues:
1. **Error: No Images Found**
   - Ensure the folder contains images in supported formats.
   - Verify the folder path entered during the macro execution.

2. **Image Placement Issues**
   - Check that your Excel sheet is empty before running the macro.
   - Ensure your Excel version supports VBA macros.

3. **Macro Not Running**
   - Enable macros in Excel under `File > Options > Trust Center > Trust Center Settings > Macro Settings`.

---

## Contribution
We welcome contributions to enhance this project! If you have suggestions, bug fixes, or feature requests, feel free to:
1. Fork the repository.
2. Create a new branch for your changes.
3. Submit a pull request with a detailed description of your changes.

---

## Contact Me
If you have any questions, feedback, or collaboration ideas, feel free to reach out:
- **Email:** [jordan.c.l.wright@gmail.com](mailto:jordan.c.l.wright@gmail.com)

---

## License
This project is licensed under the [MIT License](LICENSE), allowing free use, modification, and distribution. Please include proper attribution when sharing.

---

Thank you for using the **Excel Image Importer with VBA**! 
