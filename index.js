const express = require('express');
const app = express();

const courses = [
    {id: 1, name: 'course1'},
    {id: 2, name: 'course2'},
    {id: 3, name: 'course3'}
]

app.get('/', (req, res) => {
    res.send('Hello World');
});

app.get('/api/courses', (req, res) => {
    res.send([1,2,3]);
})

app.get('/api/courses/:id', (req, res) => {
    let course = courses.find(c => c.id === parseInt(req.params.id));
    if (!course) res.status(404).send('Course not found');
    res.send(course)
    // res.send(req.query)
})

const port = process.env.PORT || 3000
app.listen(3000, () => console.log(`Listening on port ${port}`))