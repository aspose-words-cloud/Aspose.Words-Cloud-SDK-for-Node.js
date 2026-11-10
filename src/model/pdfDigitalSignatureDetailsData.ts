/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="pdfDigitalSignatureDetailsData.ts">
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

export const importsMapPdfDigitalSignatureDetailsData = {
};

/**
 * Container class for details of digital signature.
 */
export class PdfDigitalSignatureDetailsData implements ModelInterface {
    /**
     * Attribute type map
     */
    public static attributeTypeMap: Array<AttributeInfo> = [
        {
            name: "certificateFilename",
            baseName: "CertificateFilename",
            type: "string",
        },
        {
            name: "hashAlgorithm",
            baseName: "HashAlgorithm",
            type: "PdfDigitalSignatureDetailsData.HashAlgorithmEnum",
        },
        {
            name: "location",
            baseName: "Location",
            type: "string",
        },
        {
            name: "reason",
            baseName: "Reason",
            type: "string",
        },
        {
            name: "signatureDate",
            baseName: "SignatureDate",
            type: "Date",
        }
    ];

    /**
     * Returns attribute type map
     */
    public static getAttributeTypeMap() {
        return PdfDigitalSignatureDetailsData.attributeTypeMap;
    }

    /**
     * Gets or sets the certificate's filename using for signing.
     */
    public certificateFilename: string;

    /**
     * Gets or sets the hash algorithm.
     */
    public hashAlgorithm: PdfDigitalSignatureDetailsData.HashAlgorithmEnum;

    /**
     * Gets or sets the location of the signing.
     */
    public location: string;

    /**
     * Gets or sets the reason for the signing.
     */
    public reason: string;

    /**
     * Gets or sets the date of the signing.
     */
    public signatureDate: Date;

    public constructor(init?: Partial< PdfDigitalSignatureDetailsData >) {
        Object.assign(this, init);
    }

    public collectFilesContent(_resultFilesContent: Array<any>) {
    }

    public validate() {
    }
}

/**
 * Enums for PdfDigitalSignatureDetailsData
 */
// tslint:disable:quotemark
// tslint:disable-next-line:no-namespace
export namespace PdfDigitalSignatureDetailsData {
    export enum HashAlgorithmEnum {
        Sha256 = 'Sha256' as any,
        Sha384 = 'Sha384' as any,
        Sha512 = 'Sha512' as any,
        RipeMD160 = 'RipeMD160' as any,
        Auto = 'Auto' as any
    }
}
// tslint:enable:quotemark

