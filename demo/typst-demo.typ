#import "@preview/touying:0.5.5": *
#import themes.metropolis: *

// #import "@preview/numbly:0.1.0": numbly

#show: metropolis-theme.with(
  aspect-ratio: "16-9",
  footer: self => self.info.institution,
  config-info(
    title: [#text(weight: "bold")[`pdf2ppt`]: Convert PDF Slides to PPT],
    author: [Teddy van Jerry (Wuqiong Zhao)],
    date: datetime.today().display("[month repr:long] [day], [year]"),
  ),
)

#set text(font: "Fira Sans", weight: "light", size: 20pt)
#show math.equation: set text(font: "Fira Math")
#set strong(delta: 100)

#title-slide()

== Motivation
Typst users are often confronted with the annoying situation: \
#h(2em) _*Submit PPT files not PDF files!*_

So...
we still use `touying` to creaate the slides,
but convert the PDF slides to PPT files, each slide as a SVG converted from PDF.

// \href{https://github.com/ashafaei/pdf2pptx}{\texttt{ashafaei/pdf2pptx}}
#link("https://github.com/ashafaei/pdf2pptx")[ashafaei/pdf2pptx] is a wonderful project,
but the image quality is limited, and seems a bit buggy.

#focus-slide[
  Source Code: \
  #link("https://github.com/Teddy-van-Jerry/pdf2ppt")
]

