import urllib.request
import os

model_dir = os.path.join(os.path.dirname(__file__), '../../models/language_detection/fasttext_lid')
os.makedirs(model_dir, exist_ok=True)
model_path = os.path.join(model_dir, 'lid.176.bin')

url = "https://dl.fbaipublicfiles.com/fasttext/supervised-models/lid.176.bin"

if not os.path.exists(model_path):
    print("Downloading fastText LID model (~126MB)... this might take a minute.")
    urllib.request.urlretrieve(url, model_path)
    print("Download complete!")
else:
    print("Model already exists.")
