const screenshot=require("screenshot-desktop");
async function takeScreenshot(filename){await screenshot({filename}); return filename;}
module.exports={takeScreenshot};