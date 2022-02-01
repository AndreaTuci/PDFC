<div id="top"></div>

<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->

<!-- PROJECT LOGO -->
<br />
<div style="text-align: center" align="center">
  <a href="https://github.com/AndreaTuci/PDFC">
    <img src="img/app.png" alt="Logo" width="55" height="64">
  </a>

  <h2 style="text-align: center">PDFC</h2>

  <p style="text-align: center" align="center">
    A smart Excel to PDF converter
    <br />
    <a href="https://github.com/AndreaTuci/PDFC/issues"><strong>Request feature | Report bug »</strong></a>
    <br />
    </p>
    <a href="https://www.linkedin.com/in/andrea-tuci-065463226/" target="_blank">
        <img src="img/LinkedIn-blue.png" alt="LinkedIn" width="111" height="28">
    </a>
    <br />
</div>


<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about">About The Project</a>
    </li>
    <li>
      <a href="#installation">Installation</a>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
  </ol>
</details>


<!-- ABOUT THE PROJECT -->

<div id="about"></div>

## About The Project

<br />
<div style="text-align: center" align="center">
<img height="442" src=".\img\screenshot.png" alt="Screenshot" width="208"/>
</div>
<br />

PDFC is an office tool (**Windows only!**) that allows you to convert excel files into PDF by creating a folder for each excel file and a PDF for each sheet.
It will save you a lot of time if you have to convert many sheets often.

You can convert files by uploading them via button or simply by dragging them onto the application window, PDFC will do the rest in seconds!

<p style="text-align: right" align="right">(<a href="#top">back to top</a>)</p>

<div id="installation"></div>

## Installation

Create a virtual environment

Use [pip](https://pip.pypa.io/en/stable/) to install PDFC requirements.

```bash
# Create an environment
python3 -m venv your-new-venv

# Activate the new venv
your-new-venv\Scripts\activate

# Install requirements from requirements.txt
<your-new-venv> pip install -r requirements.txt
```

You can also create an .exe using pyinstaller or launching build.bat while into your-new-venv:

```bash
<your-new-venv> build.bat    
```

This will create a new .exe file in the dist folder, inside the project folder. 
To view the images copy the img folder next to the executable file

<p style="text-align: right" align="right">(<a href="#top">back to top</a>)</p>

<div id="usage"></div>

## Usage

Launch PDFC.py or open the .exe

```bash
<your-new-venv> python PDFC.py
```

<p style="text-align: right" align="right">(<a href="#top">back to top</a>)</p>

<div id="contributing"></div>

## Contributing

Pull requests are welcome. 

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p style="text-align: right" align="right">(<a href="#top">back to top</a>)</p>

<div id="license"></div>

## License

[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)

<p style="text-align: right" align="right">(<a href="#top">back to top</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://github.com/AndreaTuci/PDFC
[contributors-url]: https://github.com/AndreaTuci/PDFC/graphs/contributors
[forks-shield]: https://github.com/AndreaTuci/PDFC
[forks-url]: https://github.com/AndreaTuci/PDFC
[stars-shield]: https://github.com/AndreaTuci/PDFC
[stars-url]: https://github.com/AndreaTuci/PDFC
[issues-shield]: https://github.com/AndreaTuci/PDFC
[issues-url]: https://github.com/AndreaTuci/PDFC
[license-shield]: https://github.com/AndreaTuci/PDFC
[license-url]: https://github.com/AndreaTuci/PDFC
[linkedin-shield]: https://github.com/AndreaTuci/PDFC
[linkedin-url]: https://github.com/AndreaTuci/PDFC
