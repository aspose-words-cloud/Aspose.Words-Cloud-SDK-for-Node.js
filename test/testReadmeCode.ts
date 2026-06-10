/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="testReadmeCode.ts">
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

import "mocha";
import { WordsApi, UploadFileRequest, SaveAsRequest, PdfSaveOptionsData } from "../src/api";
import * as BaseTest from "./baseTest";
import { fail } from "assert";
import * as fs from "fs";

describe("Readme code tests", () => {
    it("Readme code should run without errors'", () => {

        // set your credentials here
        const config = require("../testConfig.json");
        const clientId = config.ClientId;
        const clientSecret = config.ClientSecret;
        const baseUrl = config.BaseUrl;

        const localPath = BaseTest.localCommonTestDataFolder + "test_multi_pages.docx";

        // Start README example

        // create API instance (baseUrl is optional)
        const wordsApi = new WordsApi(clientId, clientSecret, baseUrl);

        // upload file to the Aspose cloud
        const uploadRequest = new UploadFileRequest();
        uploadRequest.path = "uploaded.docx";
        uploadRequest.fileContent = fs.createReadStream(localPath);

        wordsApi.uploadFile(uploadRequest)
            .then((_uploadResult) => {
                // save the file as pdf in the cloud
                const request = new SaveAsRequest({
                    name: "uploaded.docx",
                    saveOptionsData: new PdfSaveOptionsData(
                        {
                            fileName: "destination.pdf"
                        })
                });

                wordsApi.saveAs(request)
                    .then((_result) => {
                        // deal with the pdf file
                    })
                    .catch(function(_err) {
                        // handle saveAs request error
                    });
            })
            .catch(function(_err) {
                // handle uploadFile request error
            });

        // End README example

        writeToReadme();
    });
});

const writeToReadme = () => {
    // set regex strings
    const startPatern = new RegExp("^\\s*// Start README example\\s*$");
    const endPattern = new RegExp("^\\s*// End README example\\s*$");

    // set paths
    const sourcePath = __dirname + "/testReadmeCode.ts";
    const readmePath = __dirname + "/../README.md";

    // read the file
    const codeLines = fs.readFileSync(sourcePath).toString().split("\n");

    // extract readme code
    const readmeCode = new Array();
    let skipMode = true;

    for (const line of codeLines) {
        if (skipMode)
        {
            skipMode = !startPatern.test(line);
        }

        if (!skipMode)
        {
            readmeCode.push(line);
            skipMode = endPattern.test(line);
        }
    }

    if (readmeCode.length < 2) {
        fail("No readme code found in testReadCode.ts file");
    }

    // read readme.md
    const readmeLines = fs.readFileSync(readmePath).toString().split("\n");

    // replace readme code
    const newReadmeLines = new Array();
    let codeMode = false;

    for (const line of readmeLines)
    {
        if (!codeMode)
        {
            codeMode = startPatern.test(line);

            if (codeMode)
            {
                for (const codeLine of readmeCode) newReadmeLines.push(codeLine);
            }
        }

        if (codeMode)
        {
            codeMode = !endPattern.test(line);
            continue;
        }

        if(!codeMode)
        {
            newReadmeLines.push(line);
        }
    }

    // write new content to readme
    fs.writeFileSync(readmePath, newReadmeLines.join('\n'));
}
