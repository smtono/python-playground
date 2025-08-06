from PIL import Image, ImageDraw
from textblob import TextBlob
import random

# Emotion palettes
emotion_palettes = {
    "positive": ["#ffadad", "#ffd6a5", "#fdffb6", "#caffbf"],
    "neutral": ["#d3d3d3", "#b0b0b0", "#a0a0a0", "#909090"],
    "negative": ["#a0c4ff", "#bdb2ff", "#ffc6ff", "#ffb5a7"]
}

# Analyze sentiment
def get_emotion(text):
    analysis = TextBlob(text)
    polarity = analysis.sentiment.polarity
    if polarity > 0.2:
        return "positive"
    elif polarity < -0.2:
        return "negative"
    else:
        return "neutral"

# Generate image
def generate_emotion_image(text, width=256, height=256, filename="emotion_image.png"):
    emotion = get_emotion(text)
    palette = emotion_palettes[emotion]
    
    img = Image.new("RGB", (width, height), "#000000")
    draw = ImageDraw.Draw(img)

    for x in range(width):
        for y in range(height):
            color = random.choice(palette)
            draw.point((x, y), fill=color)

    img.save(filename)
    return filename

# Example usage
journal_entry = "I'm feeling a bit off today... like I'm floating through molasses. But there's still hope."
generate_emotion_image(journal_entry)
