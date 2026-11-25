/*
 * --------------------------------------------------------------------------------
 * <copyright company="Aspose" file="drawingObjectsTests.ts">
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

// Example of how to get drawing objects.
describe("drawingObjects", () => {
    expect(fs);
    const remoteDataFolder = BaseTest.remoteBaseTestDataFolder + "/DocumentElements/DrawingObjectss";
    const localFile = "Common/test_multi_pages.docx";
    const localDrawingFile = "DocumentElements/DrawingObjects/sample_EmbeddedOLE.docx";

    // Test for getting drawing objects from document.
    describe("getDocumentDrawingObjects test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjects.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectsRequest({
                    name: remoteFileName,
                    nodePath: "sections/0",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjects(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing objects from document online.
    describe("getDocumentDrawingObjectsOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const request = new model.GetDocumentDrawingObjectsOnlineRequest({
                document: requestDocument,
                nodePath: "sections/0"
            });

            // Act
            return wordsApi.getDocumentDrawingObjectsOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for getting drawing objects from document without node path.
    describe("getDocumentDrawingObjectsWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectsWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectsRequest({
                    name: remoteFileName,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjects(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object by specified index.
    describe("getDocumentDrawingObjectByIndex test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectByIndex.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectByIndexRequest({
                    name: remoteFileName,
                    index: 0,
                    nodePath: "sections/0",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectByIndex(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object by specified index online.
    describe("getDocumentDrawingObjectByIndexOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const request = new model.GetDocumentDrawingObjectByIndexOnlineRequest({
                document: requestDocument,
                index: 0,
                nodePath: "sections/0"
            });

            // Act
            return wordsApi.getDocumentDrawingObjectByIndexOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for getting drawing object by specified index without node path.
    describe("getDocumentDrawingObjectByIndexWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectByIndexWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectByIndexRequest({
                    name: remoteFileName,
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectByIndex(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object by specified index and format.
    describe("renderDrawingObject test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectByIndexWithFormat.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.RenderDrawingObjectRequest({
                    name: remoteFileName,
                    format: "png",
                    index: 0,
                    nodePath: "sections/0",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.renderDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object by specified index and format online.
    describe("renderDrawingObjectOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const request = new model.RenderDrawingObjectOnlineRequest({
                document: requestDocument,
                format: "png",
                index: 0,
                nodePath: "sections/0"
            });

            // Act
            return wordsApi.renderDrawingObjectOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for getting drawing object by specified index and format without node path.
    describe("renderDrawingObjectWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectByIndexWithFormatWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.RenderDrawingObjectRequest({
                    name: remoteFileName,
                    format: "png",
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.renderDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for reading drawing object's image data.
    describe("getDocumentDrawingObjectImageData test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectImageData.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectImageDataRequest({
                    name: remoteFileName,
                    index: 0,
                    nodePath: "sections/0",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectImageData(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for reading drawing object's image data online.
    describe("getDocumentDrawingObjectImageDataOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const request = new model.GetDocumentDrawingObjectImageDataOnlineRequest({
                document: requestDocument,
                index: 0,
                nodePath: "sections/0"
            });

            // Act
            return wordsApi.getDocumentDrawingObjectImageDataOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for reading drawing object's image data without node path.
    describe("getDocumentDrawingObjectImageDataWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectImageDataWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectImageDataRequest({
                    name: remoteFileName,
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectImageData(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object OLE data.
    describe("getDocumentDrawingObjectOleData test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectOleData.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localDrawingFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectOleDataRequest({
                    name: remoteFileName,
                    index: 0,
                    nodePath: "sections/0",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectOleData(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for getting drawing object OLE data online.
    describe("getDocumentDrawingObjectOleDataOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localDrawingFile);
            const request = new model.GetDocumentDrawingObjectOleDataOnlineRequest({
                document: requestDocument,
                index: 0,
                nodePath: "sections/0"
            });

            // Act
            return wordsApi.getDocumentDrawingObjectOleDataOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for getting drawing object OLE data without node path.
    describe("getDocumentDrawingObjectOleDataWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestGetDocumentDrawingObjectOleDataWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localDrawingFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.GetDocumentDrawingObjectOleDataRequest({
                    name: remoteFileName,
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.getDocumentDrawingObjectOleData(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for adding drawing object.
    describe("insertDrawingObject test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestInsetDrawingObject.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const requestDrawingObject = new model.DrawingObjectInsert({
                    height: 0,
                    left: 0,
                    top: 0,
                    width: 0,
                    relativeHorizontalPosition: model.DrawingObjectInsert.RelativeHorizontalPositionEnum.Margin,
                    relativeVerticalPosition: model.DrawingObjectInsert.RelativeVerticalPositionEnum.Margin,
                    wrapType: model.DrawingObjectInsert.WrapTypeEnum.Inline
                })
                const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
                const request = new model.InsertDrawingObjectRequest({
                    name: remoteFileName,
                    drawingObject: requestDrawingObject,
                    imageFile: requestImageFile,
                    nodePath: "",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.insertDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for adding drawing object online.
    describe("insertDrawingObjectOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const requestDrawingObject = new model.DrawingObjectInsert({
                height: 0,
                left: 0,
                top: 0,
                width: 0,
                relativeHorizontalPosition: model.DrawingObjectInsert.RelativeHorizontalPositionEnum.Margin,
                relativeVerticalPosition: model.DrawingObjectInsert.RelativeVerticalPositionEnum.Margin,
                wrapType: model.DrawingObjectInsert.WrapTypeEnum.Inline
            })
            const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
            const request = new model.InsertDrawingObjectOnlineRequest({
                document: requestDocument,
                drawingObject: requestDrawingObject,
                imageFile: requestImageFile,
                nodePath: ""
            });

            // Act
            return wordsApi.insertDrawingObjectOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for adding drawing object without node path.
    describe("insertDrawingObjectWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestInsetDrawingObjectWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const requestDrawingObject = new model.DrawingObjectInsert({
                    height: 0,
                    left: 0,
                    top: 0,
                    width: 0,
                    relativeHorizontalPosition: model.DrawingObjectInsert.RelativeHorizontalPositionEnum.Margin,
                    relativeVerticalPosition: model.DrawingObjectInsert.RelativeVerticalPositionEnum.Margin,
                    wrapType: model.DrawingObjectInsert.WrapTypeEnum.Inline
                })
                const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
                const request = new model.InsertDrawingObjectRequest({
                    name: remoteFileName,
                    drawingObject: requestDrawingObject,
                    imageFile: requestImageFile,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.insertDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for adding drawing object with the image from the storage.
    describe("insertDrawingObjectByReference test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestInsertDrawingObjectByReference.docx";
            const remoteImagePath = remoteDataFolder + "/TestInsertDrawingObjectByReference.png";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                return wordsApi.uploadFileToStorage(
                    remoteImagePath,
                    BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png"
                ).then((result1) => {
                    expect(result1.response.statusMessage).to.equal("OK");
                    const requestDrawingObjectImageReference = model.FileReference.fromRemoteFilePath(remoteImagePath);
                    const requestDrawingObject = new model.DrawingObjectInsertData({
                        height: 0,
                        left: 0,
                        top: 0,
                        width: 0,
                        relativeHorizontalPosition: model.DrawingObjectInsertData.RelativeHorizontalPositionEnum.Margin,
                        relativeVerticalPosition: model.DrawingObjectInsertData.RelativeVerticalPositionEnum.Margin,
                        wrapType: model.DrawingObjectInsertData.WrapTypeEnum.Inline,
                        imageReference: requestDrawingObjectImageReference
                    })
                    const request = new model.InsertDrawingObjectByReferenceRequest({
                        name: remoteFileName,
                        drawingObject: requestDrawingObject,
                        nodePath: "",
                        folder: remoteDataFolder
                    });

                    // Act
                    return wordsApi.insertDrawingObjectByReference(request)
                    .then((resultApi) => {
                        // Assert
                        expect(resultApi.response.statusCode).to.equal(200);
                    });

                });

            });

       });
    });

    // Test for adding drawing object online with the image from the request.
    describe("insertDrawingObjectByReferenceOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const requestDrawingObjectImageReferenceStream = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
            const requestDrawingObjectImageReference = model.FileReference.fromLocalFileContent(requestDrawingObjectImageReferenceStream);
            const requestDrawingObject = new model.DrawingObjectInsertData({
                height: 0,
                left: 0,
                top: 0,
                width: 0,
                relativeHorizontalPosition: model.DrawingObjectInsertData.RelativeHorizontalPositionEnum.Margin,
                relativeVerticalPosition: model.DrawingObjectInsertData.RelativeVerticalPositionEnum.Margin,
                wrapType: model.DrawingObjectInsertData.WrapTypeEnum.Inline,
                imageReference: requestDrawingObjectImageReference
            })
            const request = new model.InsertDrawingObjectByReferenceOnlineRequest({
                document: requestDocument,
                drawingObject: requestDrawingObject,
                nodePath: ""
            });

            // Act
            return wordsApi.insertDrawingObjectByReferenceOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for deleting drawing object.
    describe("deleteDrawingObject test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestDeleteDrawingObject.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.DeleteDrawingObjectRequest({
                    name: remoteFileName,
                    index: 0,
                    nodePath: "",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.deleteDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for deleting drawing object online.
    describe("deleteDrawingObjectOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const request = new model.DeleteDrawingObjectOnlineRequest({
                document: requestDocument,
                index: 0,
                nodePath: ""
            });

            // Act
            return wordsApi.deleteDrawingObjectOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for deleting drawing object without node path.
    describe("deleteDrawingObjectWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestDeleteDrawingObjectWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const request = new model.DeleteDrawingObjectRequest({
                    name: remoteFileName,
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.deleteDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for updating drawing object.
    describe("updateDrawingObject test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestUpdateDrawingObject.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const requestDrawingObject = new model.DrawingObjectUpdate({
                    left: 0
                })
                const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
                const request = new model.UpdateDrawingObjectRequest({
                    name: remoteFileName,
                    drawingObject: requestDrawingObject,
                    imageFile: requestImageFile,
                    index: 0,
                    nodePath: "",
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.updateDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });

    // Test for updating drawing object online.
    describe("updateDrawingObjectOnline test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const requestDocument = fs.createReadStream(BaseTest.localBaseTestDataFolder + localFile);
            const requestDrawingObject = new model.DrawingObjectUpdate({
                left: 0
            })
            const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
            const request = new model.UpdateDrawingObjectOnlineRequest({
                document: requestDocument,
                drawingObject: requestDrawingObject,
                imageFile: requestImageFile,
                index: 0,
                nodePath: ""
            });

            // Act
            return wordsApi.updateDrawingObjectOnline(request)
            .then((resultApi) => {
                // Assert
                expect(resultApi.response.statusCode).to.equal(200);
            });

       });
    });

    // Test for updating drawing object without node path.
    describe("updateDrawingObjectWithoutNodePath test", () => {
        it("should return response with code 200", () => {
            const wordsApi = BaseTest.initializeWordsApi();
            const remoteFileName = "TestUpdateDrawingObjectWithoutNodePath.docx";

            return wordsApi.uploadFileToStorage(
                remoteDataFolder + "/" + remoteFileName,
                BaseTest.localBaseTestDataFolder + localFile
            ).then((result0) => {
                expect(result0.response.statusMessage).to.equal("OK");
                const requestDrawingObject = new model.DrawingObjectUpdate({
                    left: 0
                })
                const requestImageFile = fs.createReadStream(BaseTest.localBaseTestDataFolder + "Common/aspose-cloud.png");
                const request = new model.UpdateDrawingObjectRequest({
                    name: remoteFileName,
                    drawingObject: requestDrawingObject,
                    imageFile: requestImageFile,
                    index: 0,
                    folder: remoteDataFolder
                });

                // Act
                return wordsApi.updateDrawingObject(request)
                .then((resultApi) => {
                    // Assert
                    expect(resultApi.response.statusCode).to.equal(200);
                });

            });

       });
    });
});