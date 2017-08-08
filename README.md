# PdfHandy.Net
Pdf wrapper to fill text/image/u3d into pdf acro form

Generally you have a pdf template with acro form to be filled, this project helps you to to fill text, image and u3d image to the form.
What you need to do is to define contracts for text/image/u3d with pdf action attributes and font attributes and call the method by passing the contract argument, that's that handy. Please note that your contract property name should be the exact the same as the acro form name in pdf.

Note that it is using iText7 to handle with text and image and spire.pdf to handle with u3d. If you like to apply to the commercail, please contact iText or spire.
