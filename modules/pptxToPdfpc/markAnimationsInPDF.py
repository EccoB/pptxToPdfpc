from pdfrw import PdfReader, PdfWriter
from pagelabels import PageLabels, PageLabelScheme
import json


class pageLabelsWithAnimations():
    """
    Class alters the Page-Labels in a PDF-Presentation such a way, that pdfpc
    detects animations as such, not incrementing the page-count.
    This helps for example when exporting the animations of a 
    PowerPointPresentation.
    As input, a manually written json-File, explaining where the animations are
    is currently needed.
    """
    def __init__(self,fileJSONPath,pdfWithAnimationSlides, outputPDFFilePath):
        """
        fileJSONPath: path to json-file explaining where the animations are
        pdfWithAnimationSlides: PDF, containing the presentation with
            animations
        outputPDFFilePath: output file for the pdf-presentation with the
            altered PDF-Labels
        """
        
        self.fileJSONPath = fileJSONPath
        self.pdfWithAnimationSlides = pdfWithAnimationSlides
        self.outputPDFFilePath = outputPDFFilePath

        # Loading
        self.animations=[]
        self.loadJsonFile()

        self.reader = PdfReader(self.pdfWithAnimationSlides)
        self.writer = PdfWriter()
        self.labels = PageLabels()
    
    
    def loadJsonFile(self):
        # Open the JSON file
        with open(self.fileJSONPath) as file:
            data = json.load(file)
            # Iterate through the JSON array
            # eccob
            for animation in data['animations']: 
                
                absStart = 0
                slideCount = int(animation['slides'])
                visibleSlideNb = 1
                name = animation.get('name','')

                
                if not self.animations:
                    absStart = animation['relStart']
                    visibleSlideNb = animation['relStart']
                else:
                    absStart = self.animations[-1]['absStart']+self.animations[-1]['slideCount']+animation['relStart']
                    visibleSlideNb = self.animations[-1]['visibleSlideNb']+animation['relStart']
                
                print(f"absStart:{absStart}, relStart: {animation['relStart']}, Slides: {slideCount} slideNb: {visibleSlideNb} Name: {name}")

                out={"absStart":absStart,"slideCount":slideCount,"visibleSlideNb":visibleSlideNb}
                self.animations.append(out)
                
        print(self.animations)
    
    def getAllAnimations(self):
        return self.animations
    
    def getCorrespondingPDFPages(self,pptxSlideNb):
        """
        Returns an array which PDF slides correspond to the pptxSlideNb
        First page starts at one
        """
        currentPage = 1
        pageList = []
        lastAnimation = None
        for animation in self.getAllAnimations():
            if animation['visibleSlideNb'] > pptxSlideNb:
                break
            else:
                lastAnimation = animation
        
        # We found the last animation, that corresponds to the number
        if(lastAnimation == None):
            pageList = [pptxSlideNb]
        else:
            if(lastAnimation['visibleSlideNb']==pptxSlideNb):
                # Matches exactly a slide with a transision: we copy the notes to each slide
                pageList = list(range(lastAnimation['absStart'],lastAnimation['absStart']+lastAnimation['slideCount']+1))
            else:            
                # Is not an animation
                #PDF startPage + Number of slides + remaining PDF slides as there are no animations left
                absPDFPage = lastAnimation['absStart']+lastAnimation['slideCount']+(pptxSlideNb-lastAnimation['visibleSlideNb'])
                pageList = [absPDFPage]
        return pageList
            
    def writePDFLabels(self):
        """
        call this function to finalize the pdf-labels and write the output filr
        """

        # Go through all animations that were stored, for each slide that,
        # has an animation, a new Page-label needs to be added for each consecutive 
        # animation-step. 
        for animation in self.getAllAnimations(): 
            for i in range(0,animation['slideCount']):
                newlabel = PageLabelScheme(startpage=animation['absStart']+i, # the index of the page of the PDF where the labels will start
                                style="arabic", # See options in PageLabelScheme.styles()
                                firstpagenum=animation['visibleSlideNb']) # number to attribute to the first page of this index
                self.labels.append(newlabel) # Adding our page labels to the existing ones
        
        # Add labels and store the output
        self.labels.write(self.reader)
        self.writer.trailer = self.reader
        self.writer.write(self.outputPDFFilePath)


        
        


if __name__ == "__main__":
    inputPDFFile = 'example/ExampleAnimationsToPDF.pdf'
    inputAnimationDescr = 'example/pagelabels.json'
    outputPDFWithAnimationsAndPageLabels = 'example/ExampleAnimationsToPDF_withPageLabels.pdf'
    pageLabeler = pageLabelsWithAnimations(inputAnimationDescr,inputPDFFile,outputPDFWithAnimationsAndPageLabels)
    pageLabeler.writePDFLabels()
