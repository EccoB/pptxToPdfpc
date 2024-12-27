# pptxToPdfpc
This repository holds the needed scripts to convert your PowerPoint presentation to a PDF, that is compatible to pdfpc - a presenting tool. In preserves the presenter-notes and animations.

This is toolset is made for people, who want to use pdfpc, but still want to work in PowerPoint. The needed PDF presentation can be directly created with PowerPoint. This is done by splitting the animations into separate slides. The typical issue is that in this case, neither notes nor the animations are available in pdfpc. The scripts help to convert those as well.

## Usage
* Create your PowerPoint presentation with animations and notes. Save it (as pptx)
* Split the animations in your PowerPoint presentation into separate slides, f.ex. use PPSplit for this
    * Export this as your presentation pdf
* Note the slide-number and length of each animation in a json-file (see example-folder for syntax.)
* Convert the PDF in a compatible manner
    * Clone this repository
    * Invoke the python script with the corresponding input parameters:
        * Path to your original PPTx-File, used to get the annotations
        * Path to your PDF file, where the animations are split into several pages
        * Path to your JSON, file describing where in the PPTx file the animations are and the number of slides in the pdfs of them
        * Path to the desired output pdf
### Example
```
export PYTHONPATH=./modules:$PYTHONPATH
python main.py example/ExampleAnimationsToPDF_Original.pptx example/ExampleAnimationsToPDF.pdf example/pagelabels.json example/finalPPTx_withComentsAndPageLabels.pdf
```
