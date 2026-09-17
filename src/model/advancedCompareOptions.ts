/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="advancedCompareOptions.ts">
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

export const importsMapAdvancedCompareOptions = {
};

/**
 * Allows to set advanced compare options.
 */
export class AdvancedCompareOptions implements ModelInterface {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "compareListDefinitions",
            baseName: "CompareListDefinitions",
            type: "boolean",
        },
        {
            name: "ignoreDmlUniqueId",
            baseName: "IgnoreDmlUniqueId",
            type: "boolean",
        },
        {
            name: "ignoreStoreItemId",
            baseName: "IgnoreStoreItemId",
            type: "boolean",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return AdvancedCompareOptions.attributeTypeMap;
    }

    /**
     * Gets or sets the value indicating whether list definition contents are compared instead of list definition Ids.
     * Default value is false.
     */
    public compareListDefinitions: boolean;

    /**
     * Gets or sets the value indicating whether to ignore difference in DrawingML unique Id.
     * Default value is false.
     */
    public ignoreDmlUniqueId: boolean;

    /**
     * Gets or sets the value indicating whether to ignore difference in StructuredDocumentTag store item Id.
     * Default value is false.
     */
    public ignoreStoreItemId: boolean;

    public constructor(init?: Partial< AdvancedCompareOptions >) {
        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
    }

    public validate() {
    }
}

