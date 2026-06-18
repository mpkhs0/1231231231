"""
============================================================
  🖱️  마우스 좌표 찾기 도우미
============================================================
사용법:
  1. 이 스크립트를 실행합니다.
  2. 마우스를 원하는 위치로 이동합니다.
  3. 실시간으로 좌표가 출력됩니다.
  4. [S] 키 → 현재 좌표 저장
  5. [Q] 키 → 종료 및 저장된 좌표 목록 출력
  6. [스크린샷] → 현재 화면 캡처 후 저장
============================================================
설치:
  pip install pyautogui keyboard pillow
============================================================
"""

import pyautogui
import keyboard
import time
import json
import os
from datetime import datetime
from PIL import ImageGrab

# 저장된 좌표 목록
saved_coords = []

def save_screenshot():
    """현재 화면을 스크린샷으로 저장"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"screenshot_{timestamp}.png"
    screenshot = ImageGrab.grab()
    screenshot.save(filename)
    print(f"\n  📸 스크린샷 저장: {filename}")
    return filename

def save_coord_with_label(x, y):
    """좌표를 라벨과 함께 저장"""
    label = input(f"\n  📌 좌표 ({x}, {y}) 에 이름을 붙이세요 (예: 메뉴버튼, 조회버튼): ").strip()
    if not label:
        label = f"좌표_{len(saved_coords)+1}"
    
    entry = {"label": label, "x": x, "y": y}
    saved_coords.append(entry)
    print(f"  ✅ 저장됨: [{label}] → ({x}, {y})")
    return entry

def export_to_json():
    """저장된 좌표를 JSON 파일로 내보내기"""
    if not saved_coords:
        print("\n  ⚠️  저장된 좌표가 없습니다.")
        return
    
    filename = "saved_coordinates.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(saved_coords, f, ensure_ascii=False, indent=2)
    print(f"\n  💾 좌표 JSON 저장: {filename}")

def print_coords_as_code():
    """저장된 좌표를 RPA 코드 형식으로 출력"""
    if not saved_coords:
        return
    
    print("\n" + "="*55)
    print("  📋 RPA 코드에 붙여넣을 좌표 설정")
    print("="*55)
    print("\nCOORDS = {")
    for item in saved_coords:
        print(f'    "{item["label"]}": ({item["x"]}, {item["y"]}),')
    print("}\n")

def run_finder():
    print("="*55)
    print("  🖱️  마우스 좌표 찾기 도우미 시작!")
    print("="*55)
    print("  [S]     → 현재 좌표 저장")
    print("  [P]     → 스크린샷 캡처")
    print("  [Q]     → 종료 및 좌표 내보내기")
    print("  마우스를 원하는 위치로 이동하세요")
    print("="*55 + "\n")

    last_pos = None
    
    try:
        while True:
            x, y = pyautogui.position()
            pos = (x, y)

            # 위치 변경될 때만 출력 (깜빡임 방지)
            if pos != last_pos:
                print(f"\r  현재 좌표: X={x:>5}, Y={y:>5}  |  [S]저장  [P]캡처  [Q]종료", end="", flush=True)
                last_pos = pos

            # 단축키 처리
            if keyboard.is_pressed("s"):
                save_coord_with_label(x, y)
                time.sleep(0.5)  # 연속 입력 방지

            elif keyboard.is_pressed("p"):
                save_screenshot()
                time.sleep(0.5)

            elif keyboard.is_pressed("q"):
                print("\n\n  🔚 종료합니다...")
                break

            time.sleep(0.05)

    except KeyboardInterrupt:
        print("\n\n  ⛔ 강제 종료됨")

    finally:
        # 결과 출력
        print("\n" + "="*55)
        print(f"  📌 저장된 좌표 목록 ({len(saved_coords)}개)")
        print("="*55)
        for i, item in enumerate(saved_coords, 1):
            print(f"  {i}. [{item['label']}] → X={item['x']}, Y={item['y']}")

        print_coords_as_code()
        export_to_json()
        print("  ✅ 완료! saved_coordinates.json 파일을 확인하세요.\n")

if __name__ == "__main__":
    run_finder()
