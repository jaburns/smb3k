const path = require('path');
const express = require('express');

const app = express();
const PORT = process.env.PORT || '3000';

app.use(express.static(path.join(__dirname, '../public')));
app.get('/', (req, res) => res.redirect('/index.html'));
app.listen(PORT, () => console.log('Listening on '+PORT));