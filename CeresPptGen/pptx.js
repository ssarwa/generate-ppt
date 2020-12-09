
const blob = require("./blob");
const story = require("./StoryPowerPoint");
let StoryPowerPoint = story.StoryPowerPoint;

/**
 * NOTE: to avoid "await is only valid in async function error - Azure Function
"  * REF: https://stackoverflow.com/questions/60619683/await-is-only-valid-in-async-function-error-azure-function
 *  await pptx.write : write has to be marked as await, otherwise function just returns before finishing
	.then(async (data)  : as this is inside an async function - otherwise get an error
 */
async function generatePptxSAS(context,req) {
	const storyPowerPoint = await generatePptx(context,req);
	const res = context.res;
	//Set blob name
	req.body.blobName = storyPowerPoint.filename; // blob functions expect blobname there
	
	await storyPowerPoint.pptx.write("arraybuffer") //stream()
	.catch((err) => {
		throw err;
	})
	.then(async (data) => { //NOTE: NEED THE "async" here as this is inside an async function
		await blob.uploadBlobFromBuffer(req, res, data); 

		var blobSASUrl = await blob.getBlobSASUrl(req, res);
		context.res = { status: 202,
				body:blobSASUrl 
		};
		context.done();
	})
}

async function generatePptxDownload(context,req) {
	const storyPowerPoint = await generatePptx(context,req);
	
	await storyPowerPoint.pptx.write("nodebuffer") //note that nodebuffer worked best
		.catch((err) => {
			throw err;
		})
		.then(async (data) => { //NOTE: NEED THE "async" here as this is inside an async function
			context.res = { status: 202,
					body: data,
					headers: { "Content-disposition": `attachment;filename=${storyPowerPoint.filename}`, 
					"Content-Length": data.length}
				};
			context.done();
	})
}

async function generatePptx(context,req) {
	try{
		var res = context.res;
		//const buffer = await blob.downloadBlobToBuffer(req,res); 
		var projectsString = JSON.stringify(req.body);
		var projects = JSON.parse(projectsString);
		console.log(projects.length + " Projects retrieved from Blob storage.");

		const storyPowerPoint = new StoryPowerPoint(projects);
		storyPowerPoint.generate();

		return storyPowerPoint;
	} catch (err) {
		console.log(err.message);
		return res.status(500).send({message: `Request Failed: ${err} `})

  }
}





module.exports = {
	generatePptxDownload,
	generatePptxSAS
}