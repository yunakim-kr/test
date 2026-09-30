"""
데이터 폴더(config.DATA_FOLDER)를 감시하다가, 통합 대상 파일이 정확히 3개
모이면 자동으로 parse_reports의 파싱 단계까지 실행한다.

(API 키 발급이 어려워, AI 요약(summary.md) 자동 생성 단계는 도입하지 않고
사람이 매주 Claude에게 직접 요청하는 방식을 유지한다. ai_summarize.py는
나중에 API 키가 생기면 쓸 수 있도록 남겨두되, 이 감시 스크립트에서는
호출하지 않는다.)

- 새 파일이 생기면(날짜 태그가 없어도) 우선 auto_tag_new_files()로 날짜를 붙인다.
- 그 다음 find_input_files()로 정확히 3개가 모였는지 확인한다.
  - 3개면: 원자료(raw_material.md)까지 자동 생성하고, "다음 단계(AI 요약)를
    진행하라"는 안내를 남긴다.
  - 3개가 아니면: 조용히 대기한다 (에러로 멈추지 않음 - 아직 다 안 모인
    정상적인 상태이기 때문).
- 한 번 처리를 완료한 파일 조합은 다시 자동으로 재처리하지 않는다
  (파일 목록이 바뀌어야 다시 실행됨) - 같은 3개로 계속 반복 실행되는 것을 방지.

실행 방법: python watch_folder.py  (콘솔을 열어둔 채로 계속 실행되어야 함)
종료 방법: Ctrl+C

watchdog 패키지가 있으면 파일 시스템 이벤트로 즉시 반응하고, 없으면 2초
간격 폴링으로 동작한다 (기능은 동일, watchdog가 조금 더 즉각적).
"""
import functools
import os
import time

import config
import parse_reports

# 백그라운드로 실행되면 표준출력이 파일로 리다이렉트되어 완전 버퍼링되므로,
# print를 즉시 flush하도록 강제한다 (로그가 늦게 보이거나 안 보이는 문제 방지).
print = functools.partial(print, flush=True)


def _current_signature(folder):
    """폴더 안 통합 대상 파일들의 (이름, 수정시각) 조합 - 처리 완료 여부 판단용."""
    try:
        files = parse_reports.find_input_files(folder)
    except SystemExit:
        return None
    return tuple(sorted((os.path.basename(f), os.path.getmtime(f)) for f in files))


def _generate_raw_material(folder, files):
    """실제 .hwp/.txt 읽기 + raw_material.md 생성 (비용이 드는 부분)."""
    import hwp_reader

    try:
        raw = parse_reports.build_raw_material(files)
    except hwp_reader.HwpReadError as e:
        print(f"[감시][중단] .hwp 읽기 실패: {e}")
        return

    out_path = os.path.join(folder, config.RAW_MATERIAL_FILENAME)
    parse_reports.write_raw_material(raw, out_path)
    print(f"[감시] 입력 파일 {len(files)}개 자동 파싱 완료 -> {out_path}")
    print("[감시] 다음 단계: raw_material.md를 확인하고 Claude에게 요약(summary.md)을 요청해 주세요.")


def tick(folder, last_signature):
    """폴더 상태를 한 번 점검한다.
    1) 날짜 태그 없는 새 파일은 매번 태깅 시도 (이게 먼저라야, 방금 올라온
       plain-name 파일도 바로 인식 대상에 들어온다).
    2) 인식된 파일 조합(signature)이 이전과 달라졌을 때만 실제 파싱을 돌린다
       (매 이벤트마다 .hwp를 여는 비용을 피하기 위함).
    반환값: 갱신된 last_signature."""
    renamed = parse_reports.auto_tag_new_files(folder)
    for old_name, new_name in renamed:
        print(f"[감시] 날짜 태그 자동 추가: {old_name} -> {new_name}")

    signature = _current_signature(folder)
    if signature is None:
        return "UNSET"
    if signature != last_signature:
        try:
            files = parse_reports.find_input_files(folder)
        except SystemExit as e:
            print(f"[감시] {e}")
            return "UNSET"
        _generate_raw_material(folder, files)
    return signature


def run_polling(folder=config.DATA_FOLDER, interval=2):
    print(f"[감시] 폴링 방식으로 '{folder}' 폴더를 지켜봅니다 ({interval}초 간격). 종료: Ctrl+C")
    last_signature = "UNSET"
    while True:
        last_signature = tick(folder, last_signature)
        time.sleep(interval)


def run_watchdog(folder=config.DATA_FOLDER):
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    state = {"last_signature": "UNSET"}

    class Handler(FileSystemEventHandler):
        def on_any_event(self, event):
            if event.is_directory:
                return
            try:
                time.sleep(1)  # 파일 복사가 끝날 때까지 짧게 대기(디바운스)
                state["last_signature"] = tick(folder, state["last_signature"])
            except Exception as e:
                # 여기서 예외가 새면 watchdog 내부 스레드가 조용히 멈출 수 있어
                # 반드시 잡아서 로그로 남기고 감시는 계속되게 한다.
                print(f"[감시][예외] {type(e).__name__}: {e}")

    print(f"[감시] watchdog으로 '{folder}' 폴더를 즉시 감지합니다. 종료: Ctrl+C")
    state["last_signature"] = tick(folder, state["last_signature"])  # 시작 시 이미 있는 파일도 확인

    observer = Observer()
    observer.schedule(Handler(), folder, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    os.makedirs(config.DATA_FOLDER, exist_ok=True)
    try:
        run_watchdog()
    except ImportError:
        print("[감시] watchdog 패키지가 없어 폴링 방식으로 대신 동작합니다 (pip install watchdog 권장).")
        run_polling()
