//TODO: page breaks - .pagebreak { page-break-before: always; }
/*

type: 
Element => https://developers.google.com/apps-script/reference/document/element
InlineImage => https://developers.google.com/apps-script/reference/document/inline-image
PositionedImage => https://developers.google.com/apps-script/reference/document/positioned-image
Text => https://developers.google.com/apps-script/reference/document/text

Blob => https://developers.google.com/apps-script/reference/base/blob

*/

/**
 * is value usable?
 *
 * @param {*} value
 */
function isUsable(value) {
    return (typeof value === 'undefined' !== true && value !== null);
}


// Use this code for Google Docs, Forms, or new Sheets.
function onOpen() {
    DocumentApp.getUi() // SpreadsheetApp Or DocumentApp or FormApp.
        .createMenu('Functions')
        .addItem('Email CleanHtml', 'EmailCleanHtml')
        .addItem('Email CleanHtml and InlineImages', 'EmailCleanHtmlAndInlineImages')
        .addItem('Email CleanHtml and Attachments', 'EmailCleanHtmlAndAttachments')
        .addItem('Create CleanHtml Document', 'CreateCleanHtmlDocument')
        .addItem('Create CleanHtml Document With Inline Images as Zip', 'CreateHtmlZipFileCleanHtmlAndInlineImages')
        .addItem('Create CleanHtml Document as Zip', 'CreateHtmlZipFile')
        .addItem('Append CleanHtml on Document', 'AppendCleanHtmlOnDocument')
        .addToUi();
}

function sendLog() {
    var body = Logger.getLog();
    DriveApp.createFile("cleanlog.txt", body);
}

function EmailCleanHtmlAndAttachments() {
    var data = ConvertGoogleDocToCleanHtml();
    var zip = createZip(data.html, data.images);
    emailHtml(data.html, data.images, true, zip);
	sendLog();
}

function EmailCleanHtml() {
    var data = ConvertGoogleDocToCleanHtml();
    emailHtml(data.html, data.images, false);
    sendLog();
}

function EmailCleanHtmlAndInlineImages() {
    var data = ConvertGoogleDocToCleanHtml({
        inlineImages: true
    });
    emailHtml(data.html, data.images, false);
	sendLog();
}

function CreateCleanHtmlDocument() {
    var data = ConvertGoogleDocToCleanHtml();
	
    createDocumentForHtml(data.html, data.images);
	sendLog();
}

function CreateHtmlZipFileCleanHtmlAndInlineImages() {
    var data = ConvertGoogleDocToCleanHtml({
        inlineImages: true
    });
    var zip = createZip(data.html, data.images);
    saveZipFile(zip);
	sendLog();
}

function CreateHtmlZipFile() {
    var data = ConvertGoogleDocToCleanHtml({
        relativeImagePath: true
    });
    var zip = createZip(data.html, data.images);
    saveZipFile(zip);
	sendLog();
}

function AppendCleanHtmlOnDocument() {
    var data = ConvertGoogleDocToCleanHtml(null, true);
    var body = DocumentApp.getActiveDocument().getBody();
    body.appendParagraph(data.html);
	sendLog();
}



/**
 * @param {boolean} inlineImages
 * @returns {{ html: string, images: {"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]}}
 */
function ConvertGoogleDocToCleanHtml(imagesOptions, disableFullBody) {
    var body = DocumentApp.getActiveDocument().getBody();
    var numChildren = body.getNumChildren();
    var output = [];
    var images = [];
    var listCounters = {};

    //TODO: Body:
    /*getMarginBottom()	Number	Retrieves the bottom margin, in points.
    getMarginLeft()	Number	Retrieves the left margin, in points.
    getMarginRight()	Number	Retrieves the right margin.
    getMarginTop()	Number	Retrieves the top margin.
    getPageHeight()	Number	Retrieves the page height, in points.
    getPageWidth()    Number  Retrieves the page width, in points.*/

    // Walk through all the child elements of the body.
    for (var i = 0; i < numChildren; i++) {
        var child = body.getChild(i);
        output.push(processItem(child, listCounters, images, imagesOptions));
    }

	if (!disableFullBody) {
		output.unshift("<html><head><meta charset='utf-8'><title>" + DocumentApp.getActiveDocument().getName() + "</title></head><body>");
		output.push("</body></html>");
	}
    var html = output.join('\r');

    return {
        "html": html,
        "images": images
    };
}

/**
 * @param {string} html
 * @param {boolean} addImagesAsAttachments
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 * @param {Blob} zip
 */
function emailHtml(html, images, addAttachments, zip) {
    var name = DocumentApp.getActiveDocument().getName();

    var attachments = [];
    if (addAttachments === true) {
        for (var j = 0; j < images.length; j++) {
            attachments.push({
                "fileName": images[j].name,
                "mimeType": images[j].type,
                "content": images[j].blob.getBytes()
            });
        }
    }

    if (isUsable(zip)) {
        attachments.push({
            "fileName": name + ".zip",
            "mimeType": "application/zip",
            "content": zip.getBytes()
        });
    }

    var inlineImages = {};
    for (var j = 0; j < images.length; j++) {
        inlineImages[[images[j].name]] = images[j].blob;
    }

    if (addAttachments === true) {
        attachments.push({
            "fileName": name + ".html",
            "mimeType": "text/html",
            "content": html
        });
    }

    MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: name,
        htmlBody: html,
        inlineImages: inlineImages,
        attachments: attachments
    });
}

/**
 * @param {string} html
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 * @returns {Blob}
 */
function createZip(html, images) {
    var name = DocumentApp.getActiveDocument().getName();
    var dataList = [];

    for (var j = 0; j < images.length; j++) {
        dataList.push(images[j].blob);
    }

    //var encoded = Utilities.base64Encode(html);
    //var byteDataArray = Utilities.base64Decode(encoded);
    //var blob = Utilities.newBlob(byteDataArray);

    var blob = Utilities.newBlob(html, "TEXT", name + ".html");

    dataList.push(blob);

    var zip = Utilities.zip(dataList, name + ".zip");
    return zip;
}

/**
 * @param {string} html
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 */
function createDocumentForHtml(html, images) {
    var name = DocumentApp.getActiveDocument().getName() + ".html";
    var newDoc = DocumentApp.create(name);
    newDoc.getBody().setText(html);
    for (var j = 0; j < images.length; j++)
        newDoc.getBody().appendImage(images[j].blob);
    newDoc.saveAndClose();
}

/**
 * @param {string} html
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 */
function saveZipFile(zip) {
    var thisFileId = DocumentApp.getActiveDocument().getId();
    var thisFile = DriveApp.getFileById(thisFileId);
    var parentFolders = thisFile.getParents();

    var file;
    if (parentFolders.hasNext()) {
        var parentFolder = parentFolders.next();
        var parentFolderId = parentFolder.getId();
        var dir = DriveApp.getFolderById(parentFolderId);
        file = dir.createFile(zip);
    } else {
        file = DriveApp.createFile(zip);
    }

    Logger.log('ID: %s, File size (bytes): %s', file.getId(), file.getSize());
}

/**
 * @param {Element} item
 */
function dumpAttributesOfItem(item) {
    dumpAttributes(item.getAttributes());
}

/**
 * @param {string[]} atts
 */
function dumpAttributes(atts) {
    // Log the paragraph attributes.
    for (var att in atts) {
		if (atts[att]) Logger.log(att + ":" + atts[att]);
    }
}

/**
 * @param {Element} item - https://developers.google.com/apps-script/reference/document/element
 * @param {Object} listCounters
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 * @returns {string}
 */
function processItem(item, listCounters, images, imagesOptions) {
    var output = [];
    var prefix = "",
        suffix = "";
    var style = "";

    var hasPositionedImages = false;
    if (item.getPositionedImages) {
        positionedImages = item.getPositionedImages();
        hasPositionedImages = true;
    }

    var itemType = item.getType();
    
    if (itemType === DocumentApp.ElementType.PARAGRAPH) {
        //https://developers.google.com/apps-script/reference/document/paragraph

        if (item.getNumChildren() == 0) {
            return "<br />";
        }

		var p = "";
        
		if (item.getIndentStart() != null) {
			p += "margin-left:" + item.getIndentStart() + "; ";
		} else {
		     // p += "margin-left: 0; "; // superfluous
		}
		
		// what does getIndentEnd actually do? the value is the same as in getIndentStart
		/*if (item.getIndentEnd() != null) {
			p += "margin-right:" + item.getIndentStart() + "; ";
		} else {
		     // p += "margin-right: 0; "; // superfluous
		}*/
		
		//Text Alignment
        switch (item.getAlignment()) {
            // Add a # for each heading level. No break, so we accumulate the right number.
            //case DocumentApp.HorizontalAlignment.LEFT:
            //  p += "text-align: left;"; break;
        case DocumentApp.HorizontalAlignment.CENTER:
            p += "text-align: center;";
            break;
        case DocumentApp.HorizontalAlignment.RIGHT:
            p += "text-align: right;";
            break;
        case DocumentApp.HorizontalAlignment.JUSTIFY:
            p += "text-align: justify;";
            break;
        default:
            p += "";
        }

        //TODO: getLineSpacing(line-height), getSpacingBefore(margin-top), getSpacingAfter(margin-bottom),

        //TODO: 
        //INDENT_END	    Enum	The end indentation setting in points, for paragraph elements.
        //INDENT_FIRST_LINE	Enum	The first line indentation setting in points, for paragraph elements.
        //INDENT_START	    Enum	The start indentation setting in points, for paragraph elements.

        if (p !== "") {
            style = 'style="' + p + '"';
        }

        //TODO: add DocumentApp.ParagraphHeading.TITLE, DocumentApp.ParagraphHeading.SUBTITLE

        //Heading or only paragraph
        switch (item.getHeading()) {
            // Add a # for each heading level. No break, so we accumulate the right number.
        case DocumentApp.ParagraphHeading.HEADING6:
            prefix = "<h6 " + style + ">", suffix = "</h6>";
            break;
        case DocumentApp.ParagraphHeading.HEADING5:
            prefix = "<h5 " + style + ">", suffix = "</h5>";
            break;
        case DocumentApp.ParagraphHeading.HEADING4:
            prefix = "<h4 " + style + ">", suffix = "</h4>";
            break;
        case DocumentApp.ParagraphHeading.HEADING3:
            prefix = "<h3 " + style + ">", suffix = "</h3>";
            break;
        case DocumentApp.ParagraphHeading.HEADING2:
            prefix = "<h2 " + style + ">", suffix = "</h2>";
            break;
        case DocumentApp.ParagraphHeading.HEADING1:
            prefix = "<h1 " + style + ">", suffix = "</h1>";
            break;
        default:
            prefix = "<p " + style + ">", suffix = "</p>";
        }

        var attr = item.getAttributes();

    } else if (itemType === DocumentApp.ElementType.INLINE_IMAGE) {
        processImage(item, images, output, imagesOptions);
    } else if (itemType === DocumentApp.ElementType.INLINE_DRAWING) {
        //TODO
        Logger.log("INLINE_DRAWING: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.LIST_ITEM) {
        var listItem = item;
        var gt = listItem.getGlyphType();
        var key = listItem.getListId() + '.' + listItem.getNestingLevel();
        var counter = listCounters[key] || 0;

        // First list item
        if (counter == 0) {
            // Bullet list (<ul>):
            if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
                prefix = '<ul style="margin:0;"><li>', suffix = '</li>';
                //suffix += "</ul>";

                // Ordered list (<ol>):
            } else {
                prefix = '<ol style="margin:0;"><li>', suffix = '</li>';
            }
        } else {
            prefix = "<li>";
            suffix = "</li>";
        }

        if (item.isAtDocumentEnd() || (item.getNextSibling() && (item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM))) {
            if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
                suffix += "</ul>";

                // Ordered list (<ol>):
            } else {
                suffix += "</ol>";
            }
        }

        counter++;
        listCounters[key] = counter;
    } else if (itemType === DocumentApp.ElementType.TABLE) {
        var row = item.getRow(0)
        var numCells = row.getNumCells();
        var tableWidth = 0;

        for (var i = 0; i < numCells; i++) {
            tableWidth += item.getColumnWidth(i);
        }
        Logger.log("TABLE tableWidth: " + tableWidth);

        //https://stackoverflow.com/questions/339923/set-cellpadding-and-cellspacing-in-css
        var style = ' style="border-collapse: collapse; width:' + tableWidth + 'px; "';

        prefix = '<table' + style + '>', suffix = "</table>";
        //Logger.log("TABLE: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.TABLE_ROW) {

        var minimumHeight = item.getMinimumHeight();
        Logger.log("TABLE_ROW getMinimumHeight: " + minimumHeight);

        prefix = "<tr>", suffix = "</tr>";
        //Logger.log("TABLE_ROW: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.TABLE_CELL) {
        /*
        BACKGROUND_COLOR	Enum	The background color of an element (Paragraph, Table, etc) or document.
        BORDER_COLOR	Enum	The border color, for table elements.
        BORDER_WIDTH	Enum	The border width in points, for table elements.
        PADDING_BOTTOM	Enum	The bottom padding setting in points, for table cell elements.
        PADDING_LEFT	Enum	The left padding setting in points, for table cell elements.
        PADDING_RIGHT	Enum	The right padding setting in points, for table cell elements.
        PADDING_TOP	    Enum	The top padding setting in points, for table cell elements.
        VERTICAL_ALIGNMENT	Enum	The vertical alignment setting, for table cell elements.
        WIDTH	        Enum	The width setting, for table cell and image elements.
        */

        //https://wiki.selfhtml.org/wiki/HTML/Tabellen/Zellen_verbinden
        var colSpan = item.getColSpan();
        Logger.log("TABLE_CELL getColSpan: " + colSpan);
        // colspan="3"

        var rowSpan = item.getRowSpan();
        Logger.log("TABLE_CELL getRowSpan: " + rowSpan);
        // rowspan ="3"

        //TODO: WIDTH must be reculculatet in percent
        var atts = item.getAttributes();

        var style = ' style=" width:' + atts.WIDTH + 'px; border: 1px solid black; padding: 5px;"';

        prefix = '<td' + style + '>', suffix = "</td>";
        //Logger.log("TABLE_CELL: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.FOOTNOTE) {
        //TODO
        Logger.log("FOOTNOTE: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.HORIZONTAL_RULE) {
        output.push("<hr />");
        //Logger.log("HORIZONTAL_RULE: " + JSON.stringify(item));
    } else if (itemType === DocumentApp.ElementType.UNSUPPORTED) {
        Logger.log("UNSUPPORTED: " + JSON.stringify(item));
    }

    output.push(prefix);

    if (hasPositionedImages === true) {
        processPositionedImages(positionedImages, images, output, imagesOptions);
    }

    if (item.getType() == DocumentApp.ElementType.TEXT) {
        processText(item, output);
    } else {

        if (item.getNumChildren) {
            var numChildren = item.getNumChildren();

            // Walk through all the child elements of the doc.
            for (var i = 0; i < numChildren; i++) {
                var child = item.getChild(i);
                output.push(processItem(child, listCounters, images, imagesOptions));
            }
        }

    }

    output.push(suffix);
    return output.join('');
}

//points = pixel * 72 / 96
//1em = 16px (Browser Default wert) 
//1px = 1/16 = 0.0625em 

function pointsToPixel(points) {
    return points * 96 / 72;
}

function pixelToPoints(pixel) {
    return pixel * 72 / 96;
}

function pixelToEm(pixel) {
    return pixel / 16;
}

function emToPixel(em) {
    return em * 16;
}


/**
 * @param {Text} item - https://developers.google.com/apps-script/reference/document/text
 * @param {string[]} output
 */
function processText(item, output) {
    var text = item.getText();
    var indices = item.getTextAttributeIndices();

    if (text === '\r') {
        Logger.log("\\r: ");
        return;
    }

    for (var i = 0; i < indices.length; i++) {
        var partAtts = item.getAttributes(indices[i]);
        var startPos = indices[i];
        var endPos = i + 1 < indices.length ? indices[i + 1] : text.length;
        var partText = text.substring(startPos, endPos);
		
		partText = partText.replace(new RegExp("(\r)", 'g'), "<br />\r");
        //Logger.log(partText);
        dumpAttributes(partAtts);

        //TODO if only ITALIC use: <blockquote></blockquote>

        //TODO: change html tags to css (i, strong, u)

        //css font-style:italic;
        if (partAtts.ITALIC) {
            output.push('<i>');
        }
        //css font-weight: bold;
        if (partAtts.BOLD) {
            output.push('<strong>');
        }
        //css text-decoration: underline
        if (partAtts.UNDERLINE) {
            output.push('<u>');
        }

        var style = "";
        
		// font family, color and size changes disabled
		/*if (partAtts.FONT_FAMILY) {
            style = style + 'font-family: ' + partAtts.FONT_FAMILY + '; ';
        }
        if (partAtts.FONT_SIZE) {
            var pt = partAtts.FONT_SIZE;
            var em = pixelToEm(pointsToPixel(pt));
            style = style + 'font-size: ' + pt + 'pt;  font-size: ' + em + 'em; ';
        }
        if (partAtts.FOREGROUND_COLOR) {
            style = style + 'color: ' + partAtts.FOREGROUND_COLOR + '; '; //partAtts.FOREGROUND_COLOR
        }
        if (partAtts.BACKGROUND_COLOR) {
            style = style + 'background-color: ' + partAtts.BACKGROUND_COLOR + '; ';
        }*/
        if (partAtts.STRIKETHROUGH) {
            style = style + 'text-decoration: line-through; ';
        }

        var a = item.getTextAlignment(startPos);
        if (a !== DocumentApp.TextAlignment.NORMAL && a !== null) {
            if (a === DocumentApp.TextAlignment.SUBSCRIPT) {
                style = style + 'vertical-align : sub; font-size : 60%; ';
            } else if (a === DocumentApp.TextAlignment.SUPERSCRIPT) {
                style = style + 'vertical-align : super; font-size : 60%; ';
            }
        }

        // If someone has written [xxx] and made this whole text some special font, like superscript
        // then treat it as a reference and make it superscript.
        // Unfortunately in Google Docs, there's no way to detect superscript
        if (partText.indexOf('[') == 0 && partText[partText.length - 1] == ']') {
            if (style !== "") {
                style = ' style="' + style + '"';
            }
            output.push('<sup' + style + '>' + partText + '</sup>');
        } else if (partText.trim().indexOf('http://') == 0 || partText.trim().indexOf('https://') == 0) {
            if (style !== "") {
                style = ' style="' + style + '"';
            }
            output.push('<a' + style + ' href="' + partText + '" rel="nofollow">' + partText + '</a>');
        } else if (partAtts.LINK_URL) {
            if (style !== "") {
                style = ' style="' + style + '"';
            }
            output.push('<a' + style + ' href="' + partAtts.LINK_URL + '" rel="nofollow">' + partText + '</a>');
        } else {
            if (style !== "") {
                partText = '<span style="' + style + '">' + partText + '</span>';
            }
            output.push(partText);
        }

        if (partAtts.ITALIC) {
            output.push('</i>');
        }
        if (partAtts.BOLD) {
            output.push('</strong>');
        }
        if (partAtts.UNDERLINE) {
            output.push('</u>');
        }

    }
    //}
}

/**
 * @param {InlineImage} item - https://developers.google.com/apps-script/reference/document/inline-image
 * @param {{"blob": Blob,"type": string,"name": string, "height": number, "width": number}[]} images
 * @param {string[]} output
 */
function processImage(item, images, output, imagesOptions) {
    if (isUsable(imagesOptions) === false) {
        imagesOptions = {
            inlineImages: false,
            relativeImagePath: false
        };
    }
    if (isUsable(imagesOptions.inlineImages) === false) {
        imagesOptions.inlineImages = false;
    }
    if (isUsable(imagesOptions.relativeImagePath) === false) {
        imagesOptions.relativeImagePath = false;
    }

    images = images || [];
    var blob = item.getBlob();
    var contentType = blob.getContentType();
    var extension = "";
    if (/\/png$/.test(contentType)) {
        extension = ".png";
    } else if (/\/gif$/.test(contentType)) {
        extension = ".gif";
    } else if (/\/jpe?g$/.test(contentType)) {
        extension = ".jpg";
    } else {
        throw "Unsupported image type: " + contentType;
    }
    var imagePrefix = "Image_";
    var imageCounter = images.length;
    var name = imagePrefix + imageCounter + extension;
    blob.setName(name);
    imageCounter++;

    if (imagesOptions.inlineImages === false) {
        var p = 'cid:';
        if (imagesOptions.relativeImagePath === true) {
            p = './';
        }
        output.push('<img src="' + p + name + '" height="' + item.getHeight() + '" width="' + item.getWidth() + '" />');
        images.push({
            "blob": blob,
            "type": contentType,
            "name": name,
            "height": item.getHeight(),
            "width": item.getWidth()
        });
    } else {
        var base64encoded = Utilities.base64Encode(blob.getBytes());
        output.push('<img src="data:' + contentType + ';base64,' + base64encoded + '" height="' + item.getHeight() + '" width="' + item.getWidth() + '" />');
    }

    //dumpAttributesOfItem(item);
}

/**
 * @param {PositionedImage[]} positionedImages - https://developers.google.com/apps-script/reference/document/positioned-image
 * @param {{"blob": Blob,"type": string,"name": string}[]} images
 * @param {string[]} output
 */
function processPositionedImages(positionedImages, images, output, imagesOptions) {
    //TODO:
    //https://developers.google.com/apps-script/reference/document/positioned-image
}
