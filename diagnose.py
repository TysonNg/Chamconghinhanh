
import os
import sys
import glob

# Force UTF-8 stdout
sys.stdout.reconfigure(encoding='utf-8')

# Setup paths (giả lập môi trường như app)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORTRAIT_DIR = os.path.join(BASE_DIR, "Ảnh BV")
INPUT_IMAGES_DIR = os.path.join(BASE_DIR, "input_images")

print(f"--- SYSTEM DIAGNOSTIC ---")
print(f"Base Dir: {BASE_DIR}")
print(f"Portrait Dir: {PORTRAIT_DIR}")
print(f"Input Dir: {INPUT_IMAGES_DIR}")

# 1. Check Portrait Directory
print(f"\n[1] Checking Portrait Directory (Anh BV)...")
if os.path.exists(PORTRAIT_DIR):
    files = glob.glob(os.path.join(PORTRAIT_DIR, "*.*")) + glob.glob(os.path.join(PORTRAIT_DIR, "*", "*.*"))
    print(f"✅ FOUND 'Ảnh BV' folder. Total files: {len(files)}")
    if len(files) > 0:
        print(f"   Example: {files[0]}")
    else:
        print(f"⚠️ 'Ảnh BV' folder is EMPTY!")
else:
    print(f"❌ NOT FOUND 'Ảnh BV' folder. Please check path!")

# 2. Check FaceMatcher module
print(f"\n[2] Checking FaceMatcher Module...")
try:
    from src.face_matcher import FaceMatcher
    print("✅ Import FaceMatcher SUCCESS")
    
    print("   Initializing FaceMatcher class...")
    matcher = FaceMatcher(PORTRAIT_DIR)
    print(f"✅ Init OK. Cached {len(matcher.portrait_cache)} people.")
    
    # 3. Test DeepFace import via matcher
    print(f"\n[3] Checking DeepFace (via matcher)...")
    from src.face_matcher import get_deepface
    df = get_deepface()
    if df:
        print("✅ DeepFace loaded SUCCESS!")
        print(f"   Model name configured: {matcher.model_name}")
    else:
        print("❌ DeepFace load FAILED (returned None)")

except Exception as e:
    print(f"❌ ERROR: {e}")
    import traceback
    traceback.print_exc()

print("\n--- FINISHED ---")
