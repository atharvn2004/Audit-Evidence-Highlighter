import cv2
import numpy as np

# 1. Create a White Background (Height 200, Width 800)
# 255 means "White" in RGB
img = np.ones((200, 800, 3), dtype=np.uint8) * 255

# 2. Define Text Settings
font = cv2.FONT_HERSHEY_SIMPLEX
text_1 = "CONFIDENTIAL EVIDENCE"
text_2 = "Client: Morgan Stanley Bank"
text_3 = "Date: 27-Nov-2025"

# 3. Write Text onto the Image (Color: 0,0,0 is Black)
# Arguments: Image, Text, (x, y), Font, Size, Color, Thickness
cv2.putText(img, text_1, (50, 50), font, 1, (0, 0, 255), 2) # Red header
cv2.putText(img, text_2, (50, 100), font, 1, (0, 0, 0), 2)  # Black text
cv2.putText(img, text_3, (50, 150), font, 0.8, (0, 0, 0), 2)

# 4. Save the file
filename = "Morgan_Stanley_Evidence.png"
cv2.imwrite(filename, img)
print(f"✅ Image Created: {filename}")