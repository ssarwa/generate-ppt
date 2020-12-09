//How to use Azure Storage:
//npm install @azure/storage-blob
//var Storage = require('azure-storage');


//Required modules
const { StorageSharedKeyCredential, BlobServiceClient, generateBlobSASQueryParameters, 
        BlobSASPermissions, ContainerSASPermissions } = require("@azure/storage-blob");

//For file upload & download tests
const fs = require('fs');

/**
    * REF: https://docs.microsoft.com/en-us/javascript/api/@azure/storage-blob/blobclient?view=azure-node-latest#beginCopyFromURL_string__BlobBeginCopyFromURLOptions_
    * REF: https://stackoverflow.com/questions/62644857/how-to-copy-a-blob-in-azure-to-another-container-with-node-sdk
    * @param {*} req 
    * @param {*} res 
    */

    async function copyBlob(req,res) {
      try {
          const body = req.body; //get reference to body structure
          //From
          const fromSharedKeyCredential = new StorageSharedKeyCredential(body.fromStorageAccount , body.fromStorageAccessKey);
          const fromBlobServiceClient = new BlobServiceClient(`https://${body.fromStorageAccount}.blob.core.windows.net`,
                                      fromSharedKeyCredential);
          const fromContainerClient = fromBlobServiceClient.getContainerClient(body.fromContainerName);
          
         
         // To: allow these to be empty -> means they are the same as the from
         let toSharedKeyCredential;
         let toBlobServiceClient;

         if (body.toStorageAccount) {
            toSharedKeyCredential = new StorageSharedKeyCredential(body.toStorageAccount , body.toStorageAccessKey)
            toBlobServiceClient =  new BlobServiceClient(`https://${body.toStorageAccount}.blob.core.windows.net`,
                               toSharedKeyCredential);
          } else {
            toSharedKeyCredential = fromSharedKeyCredential
            toBlobServiceClient =  fromBlobServiceClient
         }

         const toContainerClient =  (body.toContainerName) ? toBlobServiceClient.getContainerClient(body.toContainerName)
                              : toBlobServiceClient.getContainerClient(body.fromContainerName);
    
          //For SAS: example 1 day
          var d = new Date();
          var expiresOn = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
          //Generate a SAS for the origin "from" container
          const containerSAS = await generateContainerSAS(body.fromContainerName, "rwl", expiresOn, fromSharedKeyCredential);

                             
          //NOTES: these are arrays - should be same length (or toBlobNames.length== 0) or will FAIL
          const fromBlobNames = body.fromBlobNames;
          const toBlobNames = body.toBlobNames;

          for (let i = 0; i < fromBlobNames.length; i++) {
            
            const fromBlockBlobClient = fromContainerClient.getBlockBlobClient(fromBlobNames[i]);
            const toBlockBlobClient =  (toBlobNames.length) ? toContainerClient.getBlockBlobClient(toBlobNames[i]) 
                  : toContainerClient.getBlockBlobClient(fromBlobNames[i])
            
            //Example generate SAS for the "from" individual Blob - example with a days duration
            //const blobSAS = await generateBlobSAS(body.fromContainerName, fromBlobNames[i],"r", expiresOn, fromSharedKeyCredential);

            const response = await toBlockBlobClient.beginCopyFromURL(fromBlockBlobClient.url + "?" + containerSAS, {
            onProgress: function(state){
              //BlobBeginCopyFromUrlPollState: not called if returns quickly
              console.log("BlobBeginCopyFromUrlPollState: " + state);
            }});
            //Don't need to wait for each one to finish
            //const result = (await response.pollUntilDone())
            //console.log(result._response.status)
            //console.log(result.copyStatus)
         }
   
         return res.status(200).send({
                  "fromStorageAccount":body.fromStorageAccount,
                  "fromStorageAccessKey":body.fromStorageAccessKey,
                  "fromContainerName": body.fromContainerName,
                  "fromBlobNames": body.fromBlobNames,
                  "toStorageAccount": body.toStorageAccount,
                  "toStorageAccessKey": body.toStorageAccessKey,
                  "toContainerName": body.toContainerName,
                  "toBlobNames": (toBlobNames.length) ? toBlobNames : fromBlobNames
              });

    } catch (err) {
       console.log(
          `Request Failed - ${err.message}`
       );
      return res.status(500).send({message: `Request Failed: ${err.message} `})

    }
 }

 //// Generate service level SAS for a blob
   //permissions: choose from "racwd"
   //new Date(new Date().valueOf() + 86400
   async function generateBlobSAS(containerName,blobName,permissions, expiresOn, sharedKeyCredential) {
    const blobSAS = generateBlobSASQueryParameters({
       containerName, // Required
       blobName, // Required
       permissions: BlobSASPermissions.parse(permissions), // Required
       expiresOn: expiresOn, // Required. Date type
     },
     sharedKeyCredential
   ).toString();
   return blobSAS;
}

async function generateContainerSAS(containerName,permissions, expiresOn, sharedKeyCredential) {
  const containerSAS = generateBlobSASQueryParameters({
     containerName: containerName, // Required
     permissions: ContainerSASPermissions.parse(permissions), // Required
     expiresOn: expiresOn, // Required. Date type
   },
   sharedKeyCredential
 ).toString();
 return containerSAS;
}

async function listBlobs(req,res) {
  try {
    const body = req.body;
    const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
    const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                  sharedKeyCredential);
    const containerClient = blobServiceClient.getContainerClient(body.containerName);
      
    // List blobs
    var i = 1;
    var blobs = [];
    for await (const blob of containerClient.listBlobsFlat( { prefix: body.prefix })) {
        console.log(`Blob ${i++}: ${blob.name}`);
        blobs.push(blob.name);
    }
    return res.status(200).send({container: body.containerName, prefix: body.prefix, blobs: blobs });
  } catch (err) {
     console.log(err.message);
    return res.status(500).send({message: `Request Failed: ${err} `})
  }
}

async function uploadBlobFromBuffer (req, res,content) {
    try {
      //Standard params
      const body = req.body;
      const containerName = body.containerName;
      const blobName = body.blobName;
       
      const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
      const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                    sharedKeyCredential);
      const containerClient = blobServiceClient.getContainerClient(containerName);
      
      const blockBlobClient = containerClient.getBlockBlobClient(blobName);
      const uploadBlobResponse = await blockBlobClient.upload(content, Buffer.byteLength(content));

    } catch (err) {
       console.log(err.message);
       return res.status(500).send({message: `Request Failed: ${err} `})
    }
  } 

// Create a blob
//With Postman assumes: a Body of type form-data, and a key of type File, called "upload"
async function uploadBlobFromFile (req, res) {
    try {
       //File params
       const file = req.files.upload;
       const blobName = file.name;
       const content = file.data;

      //Standard params
      const body = req.body;
      const containerName = body.containerName;
      const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
      const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                    sharedKeyCredential);
      const containerClient = blobServiceClient.getContainerClient(body.containerName);
      
      const blockBlobClient = containerClient.getBlockBlobClient(blobName);
      const uploadBlobResponse = await blockBlobClient.upload(content, Buffer.byteLength(content));

       //Generate SAS - example with a days duration
       var d = new Date();
       var expiresOn = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
       const blobSAS = await generateBlobSAS(containerName, blobName,"r", expiresOn, sharedKeyCredential);
       var url = `https://${body.storageAccount}.blob.core.windows.net/${body.containerName}/${blobName}?` + blobSAS;
       console.log(`Upload block blob ${blobName} successfully. url ${url}`);
       return res.status(200).send({ containerName: containerName, blobName: blobName,url: url, length: content.length, etag: uploadBlobResponse.etag })
    } catch (err) {
       console.log(err.message);
       return res.status(500).send({message: `Request Failed: ${err} `})
    }
  } 

   async function downloadBlobToBuffer(req,res) {
    try {
     //Standard params
      const body = req.body;
      const containerName = body.containerName;
      const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
      const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                    sharedKeyCredential);
      const containerClient = blobServiceClient.getContainerClient(containerName);
      const blobName = body.blobName;
      const blockBlobClient = containerClient.getBlockBlobClient(blobName);
      const downloadBlockBlobResponse = await blockBlobClient.download(0);
      const buffer = await streamToBuffer(downloadBlockBlobResponse.readableStreamBody);
      return buffer;

    } catch (err) {
      console.log(err.message);
      return res.status(500).send({message: `Request Failed: ${err} `})

    }
 }

 async function getBlobSASUrl(req,res) {
  try {
   //Standard params
    const body = req.body;
    const containerName = body.containerName;
    const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
    const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                  sharedKeyCredential);
    const blobName = body.blobName;

    //Generate SAS - example with a days duration
    var d = new Date();
    var expiresOn = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
    const blobSAS = await generateBlobSAS(containerName, blobName,"r", expiresOn, sharedKeyCredential);
    var url = `https://${body.storageAccount}.blob.core.windows.net/${body.containerName}/${blobName}?` + blobSAS;
    //Return the details
    return { containerName: containerName, blobName: blobName, url: url};

  } catch (err) {
    console.log(err.message);
    return res.status(500).send({message: `Request Failed: ${err} `})

  }
}
  
   async function downloadBlob(req,res) {
      try {
       //Standard params
        const blobName = req.body.blobName;
        const buffer = await downloadBlobToBuffer(req,res);
        //Test - write to local disk:
        fs.writeFile(blobName, buffer, function (err) {
          if (err) throw err;
          console.log(`Downloaded blob ${ blobName } successfully`, downloadBlockBlobResponse.requestId);
          return res.status(200).send({ name: blobName , length:buffer.length })
        });

      } catch (err) {
        console.log(err.message);
        return res.status(500).send({message: `Request Failed: ${err} `})

      }
   }
     
   
   async function listContainers(req,res) {
     //Standard params
     const body = req.body;
     const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
     const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                   sharedKeyCredential);
     
    try {
        let i = 1;
        var containers = [];
        for await (const container of blobServiceClient.listContainers()) {
          console.log(`Container ${i++}: ${container.name}`);
          containers.push(container.name);
        }
        return res.status(200).send({containers: containers.join(",") });
    } catch (err) {
        console.log(err.message);
        return res.status(500).send({message: `Request Failed: ${err} `})
      }
  }
   // Create a container
   async function createContainer(req,res) {
     try {
       //Standard params
       const body = req.body;
       const containerName = body.containerName;
       const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
       const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                     sharedKeyCredential);
       const containerClient = blobServiceClient.getContainerClient(containerName);
      
         const createContainerResponse = await containerClient.create();
        console.log(`Created container ${containerName} successfully`, createContainerResponse.requestId);
        return res.status(200).send(`Created container '${containerName}' successfully` );
      } catch (err) {
          console.log(err.message);
          return res.status(500).send({message: `Request Failed: ${err} `})
        }
  }
  
  async function deleteContainer(req,res) {
    try{
      //Standard params
      const body = req.body;
      const containerName = body.containerName;
      const sharedKeyCredential = new StorageSharedKeyCredential(body.storageAccount , body.storageAccessKey);
      const blobServiceClient = new BlobServiceClient(`https://${body.storageAccount}.blob.core.windows.net`,
                                    sharedKeyCredential);
      const containerClient = blobServiceClient.getContainerClient(containerName);
     
        // Delete container
        await containerClient.delete();
        console.log("deleted container " + containerName);
        return res.status(200).send(`Deleted ${containerName} successfully`);
    } catch (err) {
         console.log(
             `Request Failed - ${err.details.requestId}, statusCode - ${err.statusCode}, errorCode - ${err.details.errorCode}`
         );
        console.log(err.message);
        return res.status(500).send({message: `Request Failed: ${err} `})
      }
   }
   // A helper method used to read a Node.js readable stream into a Buffer
  async function streamToBuffer(readableStream) {
      return new Promise((resolve, reject) => {
        const chunks = [];
        readableStream.on("data", (data) => {
          chunks.push(data instanceof Buffer ? data : Buffer.from(data));
        });
        readableStream.on("end", () => {
          resolve(Buffer.concat(chunks));
        });
        readableStream.on("error", reject);
      });
    }
   
   // Parallel uploading with BlockBlobClient.uploadFile() in Node.js runtime
    // BlockBlobClient.uploadFile() is only available in Node.js
    async function uploadFile(localFilePath) {
     try {
        await blockBlobClient.uploadFile(localFilePath, {
          blockSize: 4 * 1024 * 1024, // 4MB block size
          concurrency: 20, // 20 concurrency
          onProgress: (ev) => console.log(ev)
        });
        console.log("uploadFile succeeds");
      } catch (err) {
        console.log(err.message);
        return res.status(500).send({message: `Request Failed: ${err} `})
      }
    }
    // Parallel uploading a Readable stream with BlockBlobClient.uploadStream() in Node.js runtime
    // BlockBlobClient.uploadStream() is only available in Node.js
    async function uploadStream(localFilePath) {
      try {
          await blockBlobClient.uploadStream(fs.createReadStream(localFilePath), 4 * 1024 * 1024, 20, {
            abortSignal: AbortController.timeout(30 * 60 * 1000), // Abort uploading with timeout in 30mins
            onProgress: (ev) => console.log(ev)
          });
          console.log("uploadStream succeeds");
        } catch (err) {
          console.log(err.message);
          return res.status(500).send({message: `Request Failed: ${err} `})
        }
    }
    
    // Parallel uploading a browser File/Blob/ArrayBuffer in browsers with BlockBlobClient.uploadBrowserData()
    // Uncomment following code in browsers because BlockBlobClient.uploadBrowserData() is only available in browsers
    async function uploadBrowserFile(browserFile) {
      try {
          //const browserFile = document.getElementById("fileinput").files[0];
          await blockBlobClient.uploadBrowserData(browserFile, {
            blockSize: 4 * 1024 * 1024, // 4MB block size
            concurrency: 20, // 20 concurrency
            onProgress: ev => console.log(ev)
          });
      } catch (err) {
           console.log(err.message);
            return res.status(500).send({message: `Request Failed: ${err} `})
        }
    }
    // Parallel downloading a block blob into Node.js buffer
    // downloadToBuffer is only available in Node.js
    async function downloadFile(localFilePath) {
        const fileSize = fs.statSync(localFilePath).size;
        const buffer = Buffer.alloc(fileSize);
        try {
          await blockBlobClient.downloadToBuffer(buffer, 0, undefined, {
            abortSignal: AbortController.timeout(30 * 60 * 1000), // Abort uploading with timeout in 30mins
            blockSize: 4 * 1024 * 1024, // 4MB block size
            concurrency: 20, // 20 concurrency
            onProgress: (ev) => console.log(ev)
          });
          console.log("downloadToBuffer succeeds");
         } catch (err) {
          console.log(err.message);
          return res.status(500).send({message: `Request Failed: ${err} `})
        }
    }
  
   
  module.exports = {
    copyBlob,
    uploadBlobFromFile,
    downloadBlob,
    downloadBlobToBuffer,
    listBlobs,
    listContainers,
    createContainer,
    deleteContainer,
    uploadBlobFromBuffer,
    getBlobSASUrl 
  }
 
 
  
