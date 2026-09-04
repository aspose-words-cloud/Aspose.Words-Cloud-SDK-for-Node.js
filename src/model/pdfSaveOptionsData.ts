/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="pdfSaveOptionsData.ts">
 *   Copyright (c) 2026 Aspose.Words for Cloud
 * </copyright>
 * <summary>
 *   Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 * 
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 * 
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 * </summary>
 * --------------------------------------------------------------------------------
 */

import { AttributeInfo } from '../internal/attributeInfo';
import { ModelInterface } from './modelInterface';
import { DownsampleOptionsData } from './downsampleOptionsData';
import { FixedPageSaveOptionsData } from './fixedPageSaveOptionsData';
import { OutlineOptionsData } from './outlineOptionsData';
import { PdfDigitalSignatureDetailsData } from './pdfDigitalSignatureDetailsData';
import { PdfEncryptionDetailsData } from './pdfEncryptionDetailsData';

export const importsMapPdfSaveOptionsData = {
    DownsampleOptionsData,
    FixedPageSaveOptionsData,
    OutlineOptionsData,
    PdfDigitalSignatureDetailsData,
    PdfEncryptionDetailsData,
};

/**
 * Container class for pdf save options.
 */
export class PdfSaveOptionsData extends FixedPageSaveOptionsData {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "attachmentsEmbeddingMode",
            baseName: "AttachmentsEmbeddingMode",
            type: "PdfSaveOptionsData.AttachmentsEmbeddingModeEnum",
        },
        {
            name: "cacheBackgroundGraphics",
            baseName: "CacheBackgroundGraphics",
            type: "boolean",
        },
        {
            name: "compliance",
            baseName: "Compliance",
            type: "PdfSaveOptionsData.ComplianceEnum",
        },
        {
            name: "createNoteHyperlinks",
            baseName: "CreateNoteHyperlinks",
            type: "boolean",
        },
        {
            name: "customPropertiesExport",
            baseName: "CustomPropertiesExport",
            type: "PdfSaveOptionsData.CustomPropertiesExportEnum",
        },
        {
            name: "digitalSignatureDetails",
            baseName: "DigitalSignatureDetails",
            type: "PdfDigitalSignatureDetailsData",
        },
        {
            name: "displayDocTitle",
            baseName: "DisplayDocTitle",
            type: "boolean",
        },
        {
            name: "downsampleOptions",
            baseName: "DownsampleOptions",
            type: "DownsampleOptionsData",
        },
        {
            name: "embedAttachments",
            baseName: "EmbedAttachments",
            type: "boolean",
        },
        {
            name: "embedFullFonts",
            baseName: "EmbedFullFonts",
            type: "boolean",
        },
        {
            name: "encryptionDetails",
            baseName: "EncryptionDetails",
            type: "PdfEncryptionDetailsData",
        },
        {
            name: "exportDocumentStructure",
            baseName: "ExportDocumentStructure",
            type: "boolean",
        },
        {
            name: "exportLanguageToSpanTag",
            baseName: "ExportLanguageToSpanTag",
            type: "boolean",
        },
        {
            name: "fontEmbeddingMode",
            baseName: "FontEmbeddingMode",
            type: "PdfSaveOptionsData.FontEmbeddingModeEnum",
        },
        {
            name: "generateFormFieldScripts",
            baseName: "GenerateFormFieldScripts",
            type: "boolean",
        },
        {
            name: "headerFooterBookmarksExportMode",
            baseName: "HeaderFooterBookmarksExportMode",
            type: "PdfSaveOptionsData.HeaderFooterBookmarksExportModeEnum",
        },
        {
            name: "imageColorSpaceExportMode",
            baseName: "ImageColorSpaceExportMode",
            type: "PdfSaveOptionsData.ImageColorSpaceExportModeEnum",
        },
        {
            name: "imageCompression",
            baseName: "ImageCompression",
            type: "string",
        },
        {
            name: "interpolateImages",
            baseName: "InterpolateImages",
            type: "boolean",
        },
        {
            name: "openHyperlinksInNewWindow",
            baseName: "OpenHyperlinksInNewWindow",
            type: "boolean",
        },
        {
            name: "outlineOptions",
            baseName: "OutlineOptions",
            type: "OutlineOptionsData",
        },
        {
            name: "pageMode",
            baseName: "PageMode",
            type: "PdfSaveOptionsData.PageModeEnum",
        },
        {
            name: "preblendImages",
            baseName: "PreblendImages",
            type: "boolean",
        },
        {
            name: "preserveFormFields",
            baseName: "PreserveFormFields",
            type: "boolean",
        },
        {
            name: "renderChoiceFormFieldBorder",
            baseName: "RenderChoiceFormFieldBorder",
            type: "boolean",
        },
        {
            name: "textCompression",
            baseName: "TextCompression",
            type: "PdfSaveOptionsData.TextCompressionEnum",
        },
        {
            name: "useBookFoldPrintingSettings",
            baseName: "UseBookFoldPrintingSettings",
            type: "boolean",
        },
        {
            name: "useCoreFonts",
            baseName: "UseCoreFonts",
            type: "boolean",
        },
        {
            name: "useSdtTagAsFormFieldName",
            baseName: "UseSdtTagAsFormFieldName",
            type: "boolean",
        },
        {
            name: "zoomBehavior",
            baseName: "ZoomBehavior",
            type: "PdfSaveOptionsData.ZoomBehaviorEnum",
        },
        {
            name: "zoomFactor",
            baseName: "ZoomFactor",
            type: "number",
        },
        {
            name: "exportFloatingShapesAsInlineTag",
            baseName: "ExportFloatingShapesAsInlineTag",
            type: "boolean",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return super.getAttributeTypeMap().concat(PdfSaveOptionsData.attributeTypeMap);
    }

    /**
     * Gets or sets a value determining how attachments are embedded to the PDF document.
     * The default value is None and attachments are not embedded.
     * PDF/A-1, PDF/A-2 and regular PDF/A-4 (not PDF/A-4f) standards do not allow embedded files.
     * None value will be used automatically.
     */
    public attachmentsEmbeddingMode: PdfSaveOptionsData.AttachmentsEmbeddingModeEnum;

    /**
     * Gets or sets a value determining whether or not to cache graphics placed in document's background.
     * The default value is true and background graphics are written to the PDF document as an xObject. When the value is false background graphics are not cached. Some shapes are not supported for caching(shapes with fields, bookmarks, HRefs). Document background graphic is various shapes, charts, images placed in the footer or header,
     * well as background and border of a page.
     */
    public cacheBackgroundGraphics: boolean;

    /**
     * Gets or sets the PDF standards compliance level for output documents.
     */
    public compliance: PdfSaveOptionsData.ComplianceEnum;

    /**
     * Gets or sets a value indicating whether to convert footnote/endnote references in main text story into active hyperlinks.
     * When clicked the hyperlink will lead to the corresponding footnote/endnote.
     * The default value is false.
     */
    public createNoteHyperlinks: boolean;

    /**
     * Gets or sets the option that controls the way CustomDocumentProperties are exported to PDF file.
     * The default value is None.
     */
    public customPropertiesExport: PdfSaveOptionsData.CustomPropertiesExportEnum;

    /**
     * Gets or sets the details for signing the output PDF document.
     */
    public digitalSignatureDetails: PdfDigitalSignatureDetailsData;

    /**
     * Gets or sets a value indicating whether the window’s title bar should display the document title taken from the Title entry of the document information dictionary.
     * If false, the title bar should instead display the name of the PDF file containing the document.
     * The default value is false.
     */
    public displayDocTitle: boolean;

    /**
     * Gets or sets the downsample options.
     */
    public downsampleOptions: DownsampleOptionsData;

    /**
     * Gets or sets a value determining whether or not to embed attachments to the PDF document.
     * The default value is false and attachments are not embedded.
     * When the value is true attachments are embedded to the PDF document.
     * Embedding attachments is not supported when saving to PDF/A and PDF/UA compliance.
     * false value will be used automatically.
     * Embedding attachments is not supported when encryption is enabled. false value will be used automatically.
     */
    public embedAttachments: boolean;

    /**
     * Gets or sets a value indicating whether fonts are embedded into the resulting PDF documents.
     */
    public embedFullFonts: boolean;

    /**
     * Gets or sets the details for encrypting the output PDF document.
     */
    public encryptionDetails: PdfEncryptionDetailsData;

    /**
     * Gets or sets a value indicating whether to export document structure.
     */
    public exportDocumentStructure: boolean;

    /**
     * Gets or sets a value determining whether or not to create a "Span" tag in the document structure to export the text language.
     * The default value is false and "Lang" attribute is attached to a marked-content sequence in a page content stream.
     * When the value is true "Span" tag is created for the text with non-default language and "Lang" attribute is attached to this tag.
     * This value is ignored when Aspose.Words.Saving.PdfSaveOptions.ExportDocumentStructure is false.
     */
    public exportLanguageToSpanTag: boolean;

    /**
     * Gets or sets the font embedding mode.
     */
    public fontEmbeddingMode: PdfSaveOptionsData.FontEmbeddingModeEnum;

    /**
     * Gets or sets a value determining  whether to generate scripts that emulate specific Microsoft Word form field behavior in PDF.
     * Default is false.
     * When this option is enabled, the exporter generates PDF JavaScript actions to emulate Microsoft Word
     * form field behavior, such as date and time form fields with formatting and validation rules.When set to true, supported behavior will be exported as PDF JavaScript actions.
     * When set to false, no form field scripts will be generated.Script execution depends on the PDF viewer. Some PDF viewers might ignore scripts, restrict script execution,
     * or require the user to enable JavaScript.JavaScript actions are prohibited by PDF/A-1, PDF/A-2 and PDF/A-3 compliance.
     * The false value will be used automatically in this case.
     */
    public generateFormFieldScripts: boolean;

    /**
     * Gets or sets the option that controls how bookmarks in headers/footers are exported.
     * The default value is Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All.
     */
    public headerFooterBookmarksExportMode: PdfSaveOptionsData.HeaderFooterBookmarksExportModeEnum;

    /**
     * Gets or sets the option that controls how the color space will be selected for the images in PDF document.
     * The default value is "Auto". If "SimpleCmyk" value is specified, ImageCompression option is ignored and Flate compression is used for all images in the document.
     */
    public imageColorSpaceExportMode: PdfSaveOptionsData.ImageColorSpaceExportModeEnum;

    /**
     * Gets or sets the compression type to be used for all images in the document.
     */
    public imageCompression: string;

    /**
     * Gets or sets a value indicating whether image interpolation shall be performed by a conforming reader. When false is specified, the flag is not written to the output document and the default behavior of reader is used instead.
     * When the resolution of a source image is significantly lower than that of the output device, each source sample covers many device pixels. As a result, images can appear jaggy or blocky. These visual artifacts can be reduced by applying an image interpolation algorithm during rendering. Instead of painting all pixels covered by a source sample with the same color, image interpolation attempts to produce a smooth transition between adjacent sample values. A conforming Reader may choose to not implement this feature of PDF, or may use any specific implementation of interpolation that it wishes. The default value is false.
     */
    public interpolateImages: boolean;

    /**
     * Gets or sets a value indicating whether hyperlinks in the output Pdf document are forced to be opened in a new window (or tab) of a browser.
     */
    public openHyperlinksInNewWindow: boolean;

    /**
     * Gets or sets the outline options.
     */
    public outlineOptions: OutlineOptionsData;

    /**
     * Gets or sets the option that controls how the PDF document should be displayed when opened in the PDF reader.
     */
    public pageMode: PdfSaveOptionsData.PageModeEnum;

    /**
     * Gets or sets a value indicating whether to preblend transparent images with black background color.
     * Preblending images may improve PDF document visual appearance in Adobe Reader and remove anti-aliasing artifacts.In order to properly display preblended images, PDF viewer application must support /Matte entry in soft-mask image dictionary.
     * Also preblending images may decrease PDF rendering performance.The default value is false.
     */
    public preblendImages: boolean;

    /**
     * Gets or sets a value indicating whether to preserve Microsoft Word form fields as form fields in PDF or convert them to text.
     */
    public preserveFormFields: boolean;

    /**
     * Gets or sets a value indicating whether to render PDF choice form field border.
     * PDF choice form fields are used for export of SDT Combo Box Content Control, SDT Drop-Down List Content
     * Control and legacy Drop-Down Form Field when PreserveFormFields option is enabled.The default value is true.
     */
    public renderChoiceFormFieldBorder: boolean;

    /**
     * Gets or sets the compression type to be used for all textual content in the document.
     */
    public textCompression: PdfSaveOptionsData.TextCompressionEnum;

    /**
     * Gets or sets a value indicating whether the document should be saved using a booklet printing layout.
     */
    public useBookFoldPrintingSettings: boolean;

    /**
     * Gets or sets a value indicating whether to substitute TrueType fonts Arial, Times New Roman, Courier New and Symbol with core PDF Type 1 fonts.
     */
    public useCoreFonts: boolean;

    /**
     * Gets or sets a value indicating whether to use SDT control Tag or Id property as a name of form field in PDF.
     * The default value is false.When set to false, SDT control Id property is used as a name of form field in PDF.When set to true, SDT control Tag property is used as a name of form field in PDF.If set to true and Tag is empty, Id property will be used as a form field name.If set to true and Tag values are not unique, duplicate Tag values will be altered to build
     * unique PDF form field names.
     */
    public useSdtTagAsFormFieldName: boolean;

    /**
     * Gets or sets the option that controls what type of zoom should be applied when a document is opened with a PDF viewer.
     */
    public zoomBehavior: PdfSaveOptionsData.ZoomBehaviorEnum;

    /**
     * Gets or sets the zoom factor (in percentages) for a document.
     */
    public zoomFactor: number;

    /**
     * Gets or sets a value determining whether floating shapes are exported as inline tags in the document structure.
     * The default value is false and floating shapes will be exported as block-level tags,
     * placed after the paragraph in which they are anchored. When the value is true floating shapes will be exported as inline tags,
     * placed within the paragraph where they are anchored. This value is ignored when ExportDocumentStructure is false.
     */
    public exportFloatingShapesAsInlineTag: boolean;

    public constructor(init?: Partial< PdfSaveOptionsData >) {
        super(init);
        this.saveFormat = 'pdf';

        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
    }

    public validate() {
        super.validate();

        this.digitalSignatureDetails?.validate();




        this.downsampleOptions?.validate();





        this.encryptionDetails?.validate();












        this.outlineOptions?.validate();













    }
}

/**
 * Enums for PdfSaveOptionsData
 */
// tslint:disable:quotemark
// tslint:disable-next-line:no-namespace
export namespace PdfSaveOptionsData {
    export enum AttachmentsEmbeddingModeEnum {
        None = 'None' as any,
        Annotations = 'Annotations' as any,
        DocumentEmbeddedFiles = 'DocumentEmbeddedFiles' as any
    }

    export enum ComplianceEnum {
        Pdf17 = 'Pdf17' as any,
        Pdf20 = 'Pdf20' as any,
        PdfA1a = 'PdfA1a' as any,
        PdfA1b = 'PdfA1b' as any,
        PdfA2a = 'PdfA2a' as any,
        PdfA2u = 'PdfA2u' as any,
        PdfA3a = 'PdfA3a' as any,
        PdfA3u = 'PdfA3u' as any,
        PdfA4 = 'PdfA4' as any,
        PdfA4f = 'PdfA4f' as any,
        PdfA4Ua2 = 'PdfA4Ua2' as any,
        PdfUa1 = 'PdfUa1' as any,
        PdfUa2 = 'PdfUa2' as any
    }

    export enum CustomPropertiesExportEnum {
        None = 'None' as any,
        Standard = 'Standard' as any,
        Metadata = 'Metadata' as any
    }

    export enum FontEmbeddingModeEnum {
        EmbedAll = 'EmbedAll' as any,
        EmbedNonstandard = 'EmbedNonstandard' as any,
        EmbedNone = 'EmbedNone' as any
    }

    export enum HeaderFooterBookmarksExportModeEnum {
        None = 'None' as any,
        First = 'First' as any,
        All = 'All' as any
    }

    export enum ImageColorSpaceExportModeEnum {
        Auto = 'Auto' as any,
        SimpleCmyk = 'SimpleCmyk' as any
    }

    export enum PageModeEnum {
        UseNone = 'UseNone' as any,
        UseOutlines = 'UseOutlines' as any,
        UseThumbs = 'UseThumbs' as any,
        FullScreen = 'FullScreen' as any,
        UseOC = 'UseOC' as any,
        UseAttachments = 'UseAttachments' as any
    }

    export enum TextCompressionEnum {
        None = 'None' as any,
        Flate = 'Flate' as any
    }

    export enum ZoomBehaviorEnum {
        None = 'None' as any,
        ZoomFactor = 'ZoomFactor' as any,
        FitPage = 'FitPage' as any,
        FitWidth = 'FitWidth' as any,
        FitHeight = 'FitHeight' as any,
        FitBox = 'FitBox' as any
    }
}
// tslint:enable:quotemark

