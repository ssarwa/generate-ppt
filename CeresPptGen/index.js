const pptx = require("./pptx");

module.exports = async function (context, req) {
  try {
    await pptx.generatePptxDownload(context, req);
  } catch (err) {
    console.log(err.message);
    return res.status(500).send({ message: `Request Failed: ${err} ` });
  }
};
