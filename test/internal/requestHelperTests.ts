/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="requestHelperTests.ts">
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

import { expect } from "chai";
import "mocha";
import { Readable } from "stream";

import { bufferFormDataStreams } from "../../src/internal/requestHelper";

// Regression coverage for https://github.com/aspose-words-cloud/Aspose.Words-Cloud-SDK-for-Node.js/issues/16
// In 26.4.0 single-file uploads started routing through requestOptions.formData with
// { value: Readable, options: {...} }. The underlying `request` library cannot compute
// Content-Length for a Readable inside formData, so the multipart payload was sent empty
// and the Aspose API returned HTTP 400 "Unexpected end of Stream".
describe("bufferFormDataStreams", () => {
    it("buffers a Readable stream wrapped in { value, options } into a Buffer", async () => {
        const expectedBytes = Buffer.from("hello world");
        const formData: any = {
            fileContent: {
                value: Readable.from(expectedBytes),
                options: { contentType: "application/octet-stream", filename: "fileContent" },
            },
        };

        await bufferFormDataStreams(formData);

        expect(Buffer.isBuffer(formData.fileContent.value)).to.equal(true);
        expect((formData.fileContent.value as Buffer).equals(expectedBytes)).to.equal(true);
        expect(formData.fileContent.options).to.deep.equal({
            contentType: "application/octet-stream",
            filename: "fileContent",
        });
    });

    it("buffers streams inside an array entry (multiple files under the same key)", async () => {
        const expected1 = Buffer.from("file-one-content");
        const expected2 = Buffer.from("file-two-content");
        const formData: any = {
            files: [
                { value: Readable.from(expected1), options: { filename: "a.docx" } },
                { value: Readable.from(expected2), options: { filename: "b.docx" } },
            ],
        };

        await bufferFormDataStreams(formData);

        expect(Buffer.isBuffer(formData.files[0].value)).to.equal(true);
        expect(Buffer.isBuffer(formData.files[1].value)).to.equal(true);
        expect((formData.files[0].value as Buffer).equals(expected1)).to.equal(true);
        expect((formData.files[1].value as Buffer).equals(expected2)).to.equal(true);
    });

    it("leaves non-stream values untouched", async () => {
        const buf = Buffer.from("already-a-buffer");
        const formData: any = {
            plain: "plain string",
            file: { value: buf, options: { filename: "file" } },
        };

        await bufferFormDataStreams(formData);

        expect(formData.plain).to.equal("plain string");
        expect(formData.file.value).to.equal(buf);
    });

    it("is a no-op for null/undefined formData", async () => {
        await bufferFormDataStreams(undefined);
        await bufferFormDataStreams(null);
        // Reaching this point without throwing is the assertion.
    });

    it("propagates stream errors", async () => {
        const failing = new Readable({
            read() {
                this.emit("error", new Error("boom"));
            },
        });
        const formData: any = {
            fileContent: { value: failing, options: { filename: "f" } },
        };

        let caught: Error | undefined;
        try {
            await bufferFormDataStreams(formData);
        } catch (e) {
            caught = e as Error;
        }

        expect(caught).to.be.instanceOf(Error);
        expect(caught!.message).to.equal("boom");
    });
});
