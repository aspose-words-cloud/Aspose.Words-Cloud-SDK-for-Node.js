/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="folderTests.ts">
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

import * as fs from "fs";
import * as model from "../../src/model/model";
import * as BaseTest from "../baseTest";

// Example of how to work with folders.
describe("folder", () => {
    expect(fs);
    const remoteDataFolder = BaseTest.remoteBaseTestDataFolder + "/Storage";
    const localFile = "Common/test_multi_pages.docx";

    // Test for create folder.
    describe("createFolder test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const request = new model.CreateFolderRequest({
                path: remoteDataFolder + "/TestCreateFolder"
            });

            // Act
            return wordsApi.createFolder(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.statusCode).to.equal(200);
            });

       });
    });

    // Test for delete folder.
    describe("deleteFolder test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const testDeleteFolder = remoteDataFolder + "/TestDeleteFolder";

            return wordsApi.uploadFileToStorage(
                testDeleteFolder + "/TestDeleteFolder.docx",
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.DeleteFolderRequest({
                    path: testDeleteFolder,
                    recursive: true
                });

                // Act
                return wordsApi.deleteFolder(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for get file list of folder.
    describe("getFilesList test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const request = new model.GetFilesListRequest({
                path: remoteDataFolder
            });

            // Act
            return wordsApi.getFilesList(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
                expect(resultApi.body.value).to.exist;
            });

       });
    });

    // Test for copy folder.
    describe("copyFolder test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const folderToCopy = remoteDataFolder + "/TestCopyFolder";

            return wordsApi.uploadFileToStorage(
                folderToCopy + "Src/TestCopyFolderSrc.docx",
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.CopyFolderRequest({
                    destPath: folderToCopy + "Dest",
                    srcPath: folderToCopy + "Src"
                });

                // Act
                return wordsApi.copyFolder(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for move folder.
    describe("moveFolder test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/TestMoveFolderSrc/TestMoveFolderSrc.docx",
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.MoveFolderRequest({
                    destPath: BaseTest.remoteBaseTestOutFolder + "/TestMoveFolderDest_" + BaseTest.createRandomGUID(),
                    srcPath: remoteDataFolder + "/TestMoveFolderSrc"
                });

                // Act
                return wordsApi.moveFolder(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.statusCode).to.equal(200);
                });

            });

       });
    });
});