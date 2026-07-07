/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="configuration.ts">
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

import { IAuthentication, OAuth  } from "../internal/auth";
const defaultBasePath = "https://api.aspose.cloud";

/**
 * Aspose.Words for Cloud API configuration
 */
export class Configuration {
    /**
     * Authentication.
     */
    public authentication: IAuthentication;

    /**
     * Client ID.
     */
    public ClientId: string;

    /**
     * Client secret.
     */
    public ClientSecret: string;

    /**
     * Base Url.
     */
    public baseUrl: string = defaultBasePath;

    /**
     *  Gets or sets a value indicating whether debug mode. In debug mode all requests and responses are logged to console.
     */
    public debugMode: boolean;

    /**
     * Timeout in seconds. Default value is 300 seconds.
     */
    public timeout: number = 300;

    /**
     * RSA public key modulus as base64 string. Optional.
     */
    public modulus: string;

    /**
     * RSA public key exponent as base64 string. Optional.
     */
    public exponent: string;

    constructor(clientId: string, clientSecret: string, baseUrl?: string, debugMode?: boolean) {
        if (baseUrl) {
            this.baseUrl = baseUrl;
        }

        if (!clientId || !clientId.trim()) {
            throw new Error("appSID parameter must be non-empty string");
        }

        if (!clientSecret || !clientSecret.trim()) {
            throw new Error("ClientSecret parameter must be non-empty string");
        }

        this.ClientId = clientId;
        this.ClientSecret = clientSecret;
        this.debugMode = debugMode;

        this.authentication = new OAuth() as IAuthentication;
    }

    /**
     * Returns api base url
     */
    public getApiBaseUrl(): string {
        return this.baseUrl + "/v4.0";
    }
}
