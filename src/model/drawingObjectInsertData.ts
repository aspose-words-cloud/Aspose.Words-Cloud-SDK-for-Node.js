/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="drawingObjectInsertData.ts">
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
import { DrawingObjectInsert } from './drawingObjectInsert';
import { FileReference } from './fileReference';

export const importsMapDrawingObjectInsertData = {
    DrawingObjectInsert,
    FileReference,
};

/**
 * Drawing object element for insert with the image passed by reference.
 */
export class DrawingObjectInsertData extends DrawingObjectInsert {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "imageReference",
            baseName: "ImageReference",
            type: "FileReference",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return super.getAttributeTypeMap().concat(DrawingObjectInsertData.attributeTypeMap);
    }

    /**
     * Gets or sets the reference to the image file.
     * Source "Storage": the reference is either a storage path to the image located in the folder of the document or in one of its subfolders,
     * or an absolute HTTP(S) URL to download the image from. For online methods any storage path is allowed.
     * Source "Request": the reference is the name of the multipart request part that contains the image.
     */
    public imageReference: FileReference;

    public constructor(init?: Partial< DrawingObjectInsertData >) {
        super(init);
        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
        super.collectFilesContent(_resultFilesContent);
        if (this.imageReference)
        {
            this.imageReference.collectFilesContent(_resultFilesContent);
        }

    }

    public validate() {
        super.validate();
        if (this.imageReference === null || this.imageReference === undefined)
        {
            throw new Error('Property ImageReference in DrawingObjectInsertData is required.');
        }

        this.imageReference?.validate();

    }
}

