const express = require('express');

const app = express();
app.use(express.json())
app.use(express.urlencoded({ extended: true }));
const { startJob } = require('./googleSlides');

// Set up a webhook that listens for incoming data from the Google Form
app.get('/', (req, res) => {
  res.send("hi there")
})
const temp = "";
app.post('/generateCertificate', async (req, res) => {
  const formData = req.body;

  // Extract the relevant fields from the form data
  const name = req.body.name;
  const email = req.body.email;
  try {
    await startJob(name, email)
    res.status(200).json({ msg: "done" });
  }
  catch (error) {
    res.status(400).json({ error })
  }

})


app.listen(process.env.PORT || 3000, () => {
  console.log('Server started on port 3000');
});
