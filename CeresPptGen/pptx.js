const story = require("./StoryPowerPoint");
let StoryPowerPoint = story.StoryPowerPoint;

const ampStory = require("./ampcxstory");
let ampcxstory = ampStory.ampcxstory;

const archStory = require("./archoverview");
let archoverview = archStory.archoverview;

const businessStory = require("./busimpoverview");
let busimpoverview = businessStory.busimpoverview;

const retroStory = require("./retrospective");
let retrospective = retroStory.retrospective;

const techStory = require("./techoverview");
let techoverview = techStory.techoverview;

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
    var storyPowerPoint;
    var res = context.res;
    var projectsString = JSON.stringify(req.body);
    var project = JSON.parse(projectsString);
    if(project.StoryType == 'AMP Overview')
    {
      storyPowerPoint = new ampcxstory(project);
    }
    else if(project.StoryType == 'Business Overview')
    {
      storyPowerPoint = new busimpoverview(project);
    }
    else if(project.StoryType == 'Architecture Overview')
    {
      storyPowerPoint = new archoverview(project);
    }
    else if(project.StoryType == 'Retrospective')
    {
      storyPowerPoint = new retrospective(project);
    }
    else if(project.StoryType == 'Technical Overview')
    {
      storyPowerPoint = new techoverview(project);
    }
    else
    {
      storyPowerPoint = new StoryPowerPoint(project);
    }
    
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
