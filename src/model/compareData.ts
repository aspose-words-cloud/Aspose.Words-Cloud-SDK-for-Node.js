/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="compareData.ts">
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
import { AdvancedCompareOptions } from './advancedCompareOptions';
import { CompareOptions } from './compareOptions';
import { FileReference } from './fileReference';

export const importsMapCompareData = {
    AdvancedCompareOptions,
    CompareOptions,
    FileReference,
};

/**
 * Container class for compare documents.
 */
export class CompareData implements ModelInterface {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "advancedOptions",
            baseName: "AdvancedOptions",
            type: "AdvancedCompareOptions",
        },
        {
            name: "author",
            baseName: "Author",
            type: "string",
        },
        {
            name: "compareOptions",
            baseName: "CompareOptions",
            type: "CompareOptions",
        },
        {
            name: "comparingWithDocument",
            baseName: "ComparingWithDocument",
            type: "string",
        },
        {
            name: "dateTime",
            baseName: "DateTime",
            type: "Date",
        },
        {
            name: "fileReference",
            baseName: "FileReference",
            type: "FileReference",
        },
        {
            name: "resultDocumentFormat",
            baseName: "ResultDocumentFormat",
            type: "string",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return CompareData.attributeTypeMap;
    }

    /**
     * Gets or sets advanced compare options that might help to produce more precise comparison output.
     */
    public advancedOptions: AdvancedCompareOptions;

    /**
     * Gets or sets the initials of the author to use for revisions.
     */
    public author: string;

    /**
     * Gets or sets the compare options.
     */
    public compareOptions: CompareOptions;

    /**
     * Gets or sets the path to document to compare at the server.
     */
    public comparingWithDocument: string;

    /**
     * Gets or sets the date and time to use for revisions.
     */
    public dateTime: Date;

    /**
     * Gets or sets the file reference.
     */
    public fileReference: FileReference;

    /**
     * Gets or sets the result document format.
     */
    public resultDocumentFormat: string;

    public constructor(init?: Partial< CompareData >) {
        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
        if (this.fileReference)
        {
            this.fileReference.collectFilesContent(_resultFilesContent);
        }


    }

    public validate() {
        if (this.author === null || this.author === undefined)
        {
            throw new Error('Property Author in CompareData is required.');
        }
        if (this.fileReference === null || this.fileReference === undefined)
        {
            throw new Error('Property FileReference in CompareData is required.');
        }

        this.advancedOptions?.validate();




        this.compareOptions?.validate();





        this.fileReference?.validate();


    }
}

