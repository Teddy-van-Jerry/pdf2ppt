# pdf2ppt
Convert PDF Slides to PowerPoint Presentations (PPT)

## Motivation
- **LaTeX** users can easily convert the [`beamer`](https://ctan.org/pkg/beamer) slides in PDF to PPT.
- **Typst** users can easily convert the [`touying`](https://typst.app/universe/package/touying/) slides in PDF to PPT.

## Features
- [x] vector graph (highest resolution) in generated PPT
- [x] metadata (including title, author) conversion
- [x] auto-detection of slide size and aspect ratio

## Dependency
- **Python >= 3.9**: This project heavily relies on [`python-pptx`](https://python-pptx.readthedocs.io/). (Other required packages: `pypdf`, `tqdm`. Check [`requirements.txt`](requirements.txt))
- [**pdf2svg**](https://github.com/dawbarton/pdf2svg): used for converting PDF to SVG
- [**Inkscape**](https://inkscape.org/): used for converting SVG to EMF

## Technical Implementation
1. The first step is to create SVG from PDF via `pdf2svg`.
2. Due to the limitation of `python-pptx`, we need to convert SVG to EMF via `inkscape`.
3. Insert EMF into PPT via `python-pptx`.

> [!NOTE]
> `python3`, `pdf2svg` and `inkscape` should be in your PATH.
> Alternatively, `--pdf2svg-path` and `--inkscape-path` options can be used to specify their paths.

## Usage
### Installation
Use git to clone the repository.
```sh
git clone https://github.com/Teddy-van-Jerry/pdf2ppt.git --depth=1
```
If you only want the latest Python script, you can directly download the source file.
```sh
wget https://raw.githubusercontent.com/Teddy-van-Jerry/pdf2ppt/master/pdf2ppt
```

For non-Windows users,
use `make install` to install the script to `/usr/local/bin` (which should be in your `PATH` variable).

> [!TIP]
> Make sure you have the [dependency](#dependency) installed.

### Command Line Options
You can use `pdf2ppt -h` to view all options.

### Quick Start

> [!NOTE]
> If you have not installed `pdf2ppt` to your PATH, you need to use `./pdf2ppt` in the correct directory.

**Specifying the output file name**.
```sh
pdf2ppt input.pdf output.pptx
```

**Without specifying output file name**.
The output will be `input.pptx` under the same directory of input.
```sh
pdf2ppt input.pdf
```

**Toggle verbose mode**.
```sh
pdf2ppt input.pdf output.pptx --verbose
```

### Known Issues
#### Transparent Background
Unfortunately, elements with transparency are not supported by the project, due to limitations of the dependency.
You will receive a warning when such issues are detected, and you can copy the generated SVG manually to fix the problem.
View [#1](https://github.com/Teddy-van-Jerry/pdf2ppt/issues/1) for more details.

## License
Copyright ©️ 2023-2024 Teddy van Jerry ([Wuqiong Zhao](https://wqzhao.org)).
This project is distributed under the [MIT License](LICENSE).
