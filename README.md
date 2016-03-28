## Google Doc to clean HTML converter ##

Added to the original Project:
for letters :

- FONT_FAMILY
- FONT_SIZE
- FOREGROUND_COLOR
- BACKGROUND_COLOR
- STRIKETHROUGH
- TextAlignment.SUBSCRIPT
- TextAlignment.SUPERSCRIPT

and:
- rudimentary support for tables
- image weight and height


TODO:

- LineSpacing(line-height) in PARAGRAPH
- SpacingBefore(margin-top) in PARAGRAPH
- getSpacingAfter(margin-bottom) in PARAGRAPH
- complete Table support



 1. Open your Google Doc and go to Tools menu, select Script Editor. You
    should see a new window open with a nice code editor. 
 2. Copy and paste the code from here: [GoogleDocs2Html][1]
 3. Go to the File menu and Save the file the script as GoogleDoc2Html.
 4. Then from the Functions menu, in the document, choose option what you want: eMail, File or append on doc
 5. A popup window will appear titled, Authorization required.  
    Click continue to grant the following permissions:
    Know who you are on Google
    View your email address
    View and manage your documents in Google Drive
    Send email as you
 6. You will get an email at your Google Account containing the HTML
    output of the Google Doc with inline images or a zip in Google Drive or inline at the end of the doc.


  [1]: https://raw.githubusercontent.com/many20/GoogleDoc2Html/master/code.js
