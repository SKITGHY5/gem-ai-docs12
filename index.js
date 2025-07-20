const express = require('express');
const { Configuration, OpenAIApi } = require('openai');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

const openaiApiKey = process.env.OPENAI_API_KEY;
if (!openaiApiKey) {
    console.error('OPENAI_API_KEY environment variable not set');
    process.exit(1);
}

const configuration = new Configuration({
    apiKey: openaiApiKey,
});
const openai = new OpenAIApi(configuration);

app.get('/', (req, res) => {
    res.send('GeM AI Document Generator is running!');
});

app.get('/generate', async (req, res) => {
    const bid = req.query.bid;
    if (!bid) {
        return res.status(400).send('Please provide bid number as ?bid=GEM/2025/B/XXXX');
    }

    const prompt = `Generate a professional company profile for "S K IT SOLUTION" providing "Product Supply" services, for GeM bid number ${bid}.`;

    try {
        const completion = await openai.createChatCompletion({
            model: "gpt-4",
            messages: [{ role: "user", content: prompt }],
            max_tokens: 500,
        });

        const text = completion.data.choices[0].message.content;

        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Company Profile",
                                bold: true,
                                size: 32,
                            }),
                        ],
                        spacing: { after: 300 },
                    }),
                    new Paragraph(text),
                ],
            }],
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `Company_Profile_${bid.replace(/\\W+/g, '_')}.docx`;
        const filepath = path.join(__dirname, filename);
        fs.writeFileSync(filepath, buffer);

        res.download(filepath, filename, (err) => {
            if (err) {
                console.error('Download error:', err);
                res.status(500).send('Error downloading file');
            }
            fs.unlinkSync(filepath);
        });

    } catch (error) {
        console.error('OpenAI API error:', error);
        res.status(500).send('Error generating document');
    }
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
