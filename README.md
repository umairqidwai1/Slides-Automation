## Slides Automation

- Create slides automatically using a python script
- Uses the pptx file in templates folder to create all new powerpoints
- Uses the content.json file to populate the pptx file
- A new file is generated in the root folder when you run the program called "Generated_Presentation.pptx"

The n8n.json file is the n8n flow that I use currently, It runs when i recieve an email to my specific email "example+slides@gmail.com" and generates slides, uploads them to canva and emails back the link.
Feel free to download this n8n workflow and change it based on your needs. The code generates a pptx file, but you can upload it anywhere you want if you have api access.
