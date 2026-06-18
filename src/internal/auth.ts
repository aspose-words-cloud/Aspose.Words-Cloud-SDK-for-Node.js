/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="auth.ts">
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

import * as request from "request";
import { Configuration } from "./configuration";
import { invokeApiMethod } from "./requestHelper";

/**
 * Authentication logic for api calls
 */
export interface IAuthentication {
    /**
     * Apply authentication settings to header and query params.
     */
    applyToRequest(requestOptions: request.Options, configuration: Configuration): void;

    /**
     * Handle 401 response.
     */
    handle401response(configuration: Configuration);
}

/**
 * Implements OAuth authentication 
 */
export class OAuth implements IAuthentication {
    private accessToken: string;

     /**
      * Apply authentication settings to header and query params.
      */
    public async applyToRequest(requestOptions: request.Options, configuration: Configuration): Promise<void> {
        if (this.accessToken == null) {
            await this._requestToken(configuration);
        }

        if (requestOptions && requestOptions.headers) {
            requestOptions.headers.Authorization = "Bearer " + this.accessToken;
        }

        return Promise.resolve();
    }

    /**
     * Handle 401 response.
     */
    public async handle401response(configuration: Configuration) {
        await this._requestToken(configuration);
    }

    private async _requestToken(configuration: Configuration): Promise<void> {
        const requestOptions: request.Options = {
            method: "POST",
            json: true,
            uri: configuration.baseUrl + "/v4.0/words/connect/token",
            form: {
                grant_type: "client_credentials",
                client_id: configuration.ClientId,
                client_secret: configuration.ClientSecret,
            },
        };

        const response = await invokeApiMethod(requestOptions, configuration, true);
        this.accessToken = response.body.access_token;
        return Promise.resolve();
    }
}
