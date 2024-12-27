from pptxToPdfpc import *

import argparse
import os
import tempfile


def main():
    # Create the parser
    parser = argparse.ArgumentParser(description="Prepares a PDF, made in PowerPoint to be compatible to pdfpc.")
    
    # Input/Output
    parser.add_argument('inputPPTXFile', type=str, help="Path to the unaltered PPTX presentation.")
    parser.add_argument('inputPDFFile', type=str, help="Path to the PDF file, which contains the animations as a separate page.")
    parser.add_argument('inputAnimationDescr', type=str, help="Path to the JSON file, containing the positions of the animations in the presentation.")
    
    parser.add_argument('outputPDF', type=str, help="Path to the output PDF, containing animations and notes")
    parser.add_argument('--testOutput', action='store_true', help="Use the example input files.")

    # Parse the arguments
    args = parser.parse_args()
    
    # Basic checks
    if(args.testOutput):
        inputPDFFile = 'example/ExampleAnimationsToPDF.pdf'
        inputAnimationDescr = 'example/pagelabels.json'
        outputPDFWithAnimationsAndPageLabels = 'example/ExampleAnimationsToPDF_withPageLabels.pdf'

        inputPDFFileWithPageLabels = outputPDFWithAnimationsAndPageLabels
        inputPPTXFile = "example/ExampleAnimationsToPDF_Original.pptx"
        outputPDFFile = "example/ExampleAnimationsToPDF_WithCommentsAndAnimations.pdf"
    else:
        assert os.path.isfile(args.inputPPTXFile), 'PPTX File can not be read'
        assert os.path.isfile(args.inputPDFFile),  'PDF File can not be read'
        assert os.path.isfile(args.inputAnimationDescr),  'JSON File can not be read'

        inputPDFFile = args.inputPDFFile
        inputPPTXFile = args.inputPPTXFile
        inputAnimationDescr = args.inputAnimationDescr

        outputPDFFile = args.outputPDF

        # Temporary, intermediate results
        temp_dir = tempfile.gettempdir() # Create a temporary file path 
        outputPDFWithAnimationsAndPageLabels = os.path.join(temp_dir, 'pdfWithPageLabels.pdf')
        inputPDFFileWithPageLabels = outputPDFWithAnimationsAndPageLabels

    # Load the input
    pageLabeler = pageLabelsWithAnimations(inputAnimationDescr,inputPDFFile,outputPDFWithAnimationsAndPageLabels)
    pageLabeler.writePDFLabels()

    slideNote = slideNotes(inputPPTXFile,inputPDFFileWithPageLabels,outputPDFFile)
    slideNote.transferAnnotationsFromPPTxToPDF()
    slideNote.writeOutput()

    print(f"The final output can be found at {outputPDFFile}")


if __name__ == "__main__":
    main()




