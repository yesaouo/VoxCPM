conda create -n voxcpm python=3.10
conda activate voxcpm
pip3 install torch torchvision --index-url https://download.pytorch.org/whl/cu130
pip install voxcpm
python app.py --port 8808  # then open in browser: http://localhost:8808

# pptx
pip install pywin32