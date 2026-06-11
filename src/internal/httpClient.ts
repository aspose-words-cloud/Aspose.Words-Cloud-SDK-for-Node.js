/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="httpClient.ts">
 *   Copyright (c) 2025 Aspose.Words for Cloud
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

import axios, { AxiosRequestConfig, AxiosResponse } from "axios";
import { Configuration } from "./configuration";

/**
 * Invoke api method
 * @param url reuest url
 * @param requestOptions request parameters
 * @param config api configuration
 */
export async function executeRequest(options: AxiosRequestConfig, config: Configuration): Promise<AxiosResponse> {

    const client = axios.create({
        baseURL: config.baseUrl,
        timeout: 1000 * config.timeout,
        headers: {
        'x-aspose-client': 'nodejs sdk',
        'x-aspose-client-version': '24.12'
        }
    });

    return await client.request(options);
}
