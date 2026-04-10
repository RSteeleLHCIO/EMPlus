const serverless = require("serverless-http");
const appModule = require("./app.js");
const app = appModule.default || appModule;

module.exports.handler = serverless(app);
