# lyrics2ppt

<!-- ![ASCII Art](carbon.png) -->

```
   __         _         ___            __ 
  / /_ ______(_)______ |_  |___  ___  / /_
 / / // / __/ / __(_-</ __// _ \/ _ \/ __/
/_/\_, /_/ /_/\__/___/____/ .__/ .__/\__/ 
    /___/                  /_/  /_/         
```

# PowerPoint Generator from Text File

This project provides a script to generate a PowerPoint presentation from a text file. The script reads the text file, processes the content, and creates a PowerPoint presentation with customizable background colors, font colors, and font sizes.

## Features

- Generate PowerPoint slides from a text file.
- Customize background color, font color, and font size.
- Optionally add a background image with adjustable transparency.
- Supports UTF-8 encoding for text files.
- ANSI escape codes for colorful console output.

## Requirements

![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6.21-green.svg)
![Pillow](https://img.shields.io/badge/Pillow-8.2.0-yellow.svg)
![lxml](https://img.shields.io/badge/lxml-4.6.3-red.svg)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/Astraaaaaaa/lyrics2ppt.git
    cd lyrics2ppt
    ```

2. Install the required Python libraries:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

### Command Line

To generate a PowerPoint presentation from a text file, use the following command:

```sh
python generate_ppt_from_txt.py \
    --input input.txt \
    --output output.pptx \
    --bg-image background.jpg \
    --bg-color red \
    --font-color white \
    --font-size 48 \
    --transparency 0.5
```

### Arguments

- `--input`: Path to the input text file (default: `input.txt`).
- `--output`: Path to the output PowerPoint file (default: uses the first line of the input file as the title).
- `--bg-image`: Path to the background image file (default: none).
- `--bg-color`: Background color (default: `default`).
- `--font-color`: Font color (default: `white`).
- `--font-size`: Font size (default: `48`).
- `--transparency`: Transparency level for the background image (default: `0.5`).

### Sample Input File

Create a text file named `input.txt` with the following format:
```
Song Title

Song Lyrics 1 (1st slide)
Song Lyrics 2 (1st slide, new line)

Song Lyrics 3 (2nd slide)
Song Lyrics 4 (2nd slide)
Song Lyrics 5 (2nd slide)

Song Lyrics 6 (3rd slide)
Song Lyrics 7 (3rd slide)

... (and so on)
```
- The first line is the title of the presentation.
- Each paragraph separated by a blank line will be a new slide.
- Lines within a paragraph will be added as bullet points on the same slide.

## Maintainer

Astra Lee <astralee95@gmail.com>
