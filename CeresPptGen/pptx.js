const story = require("./StoryPowerPoint");
let StoryPowerPoint = story.StoryPowerPoint;

async function generatePptxDownload(context, req) {
  const storyPowerPoint = await generatePptx(context, req);

  await storyPowerPoint.pptx
    .write("nodebuffer") //note that nodebuffer worked best
    .catch((err) => {
      throw err;
    })
    .then(async (data) => {
      //NOTE: NEED THE "async" here as this is inside an async function
      context.res = {
        status: 202,
        body: data,
        headers: {
          "Content-disposition": `attachment;filename=${storyPowerPoint.filename}`,
          "Content-Length": data.length,
        },
      };
      context.done();
    });
}

async function generatePptx(context, req) {
  try {
    var res = context.res;
    var projectsString = JSON.stringify(req.body);
    var project = JSON.parse(projectsString);
    const storyPowerPoint = new StoryPowerPoint(project);
    storyPowerPoint.generate();
    return storyPowerPoint;
  } catch (err) {
    console.log(err.message);
    return res.status(500).send({ message: `Request Failed: ${err} ` });
  }
}

module.exports = {
  generatePptxDownload,
};
