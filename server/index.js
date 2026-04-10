import app from "./app.js";

const port = process.env.API_PORT || 5174;
app.listen(port, () => {
  console.log(`DynamoDB API listening on http://localhost:${port}`);
});
