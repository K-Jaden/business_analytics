"""매일 실행: YouTube 수집 → 데이터 품질 검증
사용법: python daily_run.py
"""
import subprocess, sys, os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
py = sys.executable

print('[ 1/2 ] YouTube 수집 시작...')
subprocess.run([py, 'scripts/collect_youtube.py'])

print('\n[ 2/2 ] 데이터 품질 검증...')
subprocess.run([py, 'scripts/validate_daily.py'])
