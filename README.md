<h1 align="center">
    PowerPoint Macros
</h1>
<p align="center">
  Format slides more quickly and effectively
</p>
<p align="center">
  <img src="https://img.shields.io/github/license/leliukhin/ppt-macros?label=license" alt="MIT license" />
</p>
<img src="https://github.com/leliukhin/ppt-macros/blob/HEAD/img/ribbon.png">


## Highlights
-   [x] Speeds up dozens of common tasks in PowerPoint
-   [x] Intuitive and consistent design for quick learning
-   [x] Ergonomic keyboard shortcuts for daily use (press Alt + C to try!)
-   [x] Does not interfere with PowerPoint's undo functionality
-   [x] IT-friendly .ppam installer
-   [x] Forever free to use

## Installation
1. Click [here](https://github.com/leliukhin/ppt-macros/raw/main/ppt-macros-1_0.ppam) to download `ppt-macros-1_0.ppam`
2. In PowerPoint, go to File > Options > Add-ins
3. In the bottom drop-down menu, select PowerPoint Add-ins and press Go
4. In the pop-up window, press Add New. Navigate to and select `ppt-macros-1_0.ppam`

## Features
1. **Conform shapes' properties** based on a reference shape
    * Size, position, rotation, yellow-handle adjustments, text margins 
    * Do this on one slide or across many
2. **Swap shapes** with one click
    * Anchor on any part of the shape (corners, midpoints and more)
3. **Resize multiple logos to the same size** while maintaining their aspect ratios
4. **Auto-select shapes on a messy slide** based on matching properties
5. With one click:
    * **Connect two shapes** with an elbow connector
    * **Add a NTD (note-to-draft) textbox** to a slide for your comments
    * **Create a grid of textboxes** with optional row and column headers
6. And more...

## Self-Signing
If your organization permits only digitally-signed macros, you will not be able to install the .ppam file directly. Check with your IT team if self-signed macros are accepted. Instructions to self-sign:
1. Locate SELFCERT.exe. This is usually at `C:\Program Files\Microsoft Office\root\Office16`
2. In the pop-up, enter a certificate name and press OK (the name is not important)
3. [Download](https://github.com/leliukhin/ppt-macros/raw/main/ppt-macros-1_0.pptm) and open `ppt-macros-1_0.pptm` (note: .pptm, not .ppam)
4. Press Alt + F11 to open the VBA Editor
5. Go to Tools > Digital Signature
6. Select the certificate you just created and press OK
7. Save `ppt-macros-1_0.pptm` as `ppt-macros-1_0.ppam` (note: .ppam)
8. Follow the installation [instructions](#installation) to install the .ppam file

## About
- Inspired by [Macabacus](https://macabacus.com/)
- "If the code is ugly and works, it isn't ugly." \- Albert Einstein, I hope