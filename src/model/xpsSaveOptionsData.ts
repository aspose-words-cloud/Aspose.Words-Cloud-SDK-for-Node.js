/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="xpsSaveOptionsData.ts">
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
import { DigitalSignatureDetails } from './digitalSignatureDetails';
import { FixedPageSaveOptionsData } from './fixedPageSaveOptionsData';
import { OutlineOptionsData } from './outlineOptionsData';

export const importsMapXpsSaveOptionsData = {
    DigitalSignatureDetails,
    FixedPageSaveOptionsData,
    OutlineOptionsData,
};

/**
 * Container class for xps save options.
 */
export class XpsSaveOptionsData extends FixedPageSaveOptionsData {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "bookmarksOutlineLevel",
            baseName: "BookmarksOutlineLevel",
            type: "number",
        },
        {
            name: "compressionLevel",
            baseName: "CompressionLevel",
            type: "XpsSaveOptionsData.CompressionLevelEnum",
        },
        {
            name: "digitalSignatureDetails",
            baseName: "DigitalSignatureDetails",
            type: "DigitalSignatureDetails",
        },
        {
            name: "headingsOutlineLevels",
            baseName: "HeadingsOutlineLevels",
            type: "number",
        },
        {
            name: "outlineOptions",
            baseName: "OutlineOptions",
            type: "OutlineOptionsData",
        },
        {
            name: "useBookFoldPrintingSettings",
            baseName: "UseBookFoldPrintingSettings",
            type: "boolean",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return super.getAttributeTypeMap().concat(XpsSaveOptionsData.attributeTypeMap);
    }

    /**
     * Gets or sets the level in the XPS document outline at which to display Word bookmarks.
     */
    public bookmarksOutlineLevel: number;

    /**
     * Gets or sets a compression level used to save document.
     * The default value is Normal.
     */
    public compressionLevel: XpsSaveOptionsData.CompressionLevelEnum;

    /**
     * Gets or sets the details for signing the output document.
     */
    public digitalSignatureDetails: DigitalSignatureDetails;

    /**
     * Gets or sets the number of heading levels (paragraphs formatted with the Heading styles) to include in the XPS document outline.
     */
    public headingsOutlineLevels: number;

    /**
     * Gets or sets the outline options.
     */
    public outlineOptions: OutlineOptionsData;

    /**
     * Gets or sets a value indicating whether the document should be saved using a booklet printing layout.
     */
    public useBookFoldPrintingSettings: boolean;

    public constructor(init?: Partial< XpsSaveOptionsData >) {
        super(init);
        this.saveFormat = 'xps';

        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
    }

    public validate() {
        super.validate();

        this.digitalSignatureDetails?.validate();




        this.outlineOptions?.validate();



    }
}

/**
 * Enums for XpsSaveOptionsData
 */
// tslint:disable:quotemark
// tslint:disable-next-line:no-namespace
export namespace XpsSaveOptionsData {
    export enum CompressionLevelEnum {
        Normal = 'Normal' as any,
        Maximum = 'Maximum' as any,
        Fast = 'Fast' as any,
        SuperFast = 'SuperFast' as any
    }
}
// tslint:enable:quotemark

