from flask import Flask, request, jsonify, send_file, render_template_string
import sqlite3
import pandas as pd
import zipfile
import io
from datetime import datetime

app = Flask(__name__)
DB_NAME = 'forecast_v4.db'

# 지역 정렬 순서 정의
REGIONS_ORDER = ['중앙', '강북', '서대문', '고양', '의정부', '남양주', '강릉', '원주']
CATEGORIES_ORDER = ['출동보안', '고ARPU', '영상보안(SP)', '시스템 보안(SP)', '영상보안(KT/비대면)', '시스템 보안(SP+KT/비대면)']

# 1. 데이터베이스 셋업
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS metadata 
                 (type TEXT, value TEXT, PRIMARY KEY(type, value))''')
    c.execute('''CREATE TABLE IF NOT EXISTS targets 
                 (region TEXT, category TEXT, new_target REAL, cancel_target REAL, 
                  PRIMARY KEY(region, category))''')
    c.execute('''CREATE TABLE IF NOT EXISTS actuals 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, region TEXT, category TEXT, 
                  new_actual_4w REAL, new_actual_close REAL, cancel_actual_4w REAL, cancel_actual_close REAL, 
                  timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)''')
    
    # 기초 데이터 채우기 (최초 1회)
    regions = ['중앙', '강북', '서대문', '고양', '의정부', '남양주', '강릉', '원주']
    categories = ['출동보안', '고ARPU', '영상보안(SP)', '시스템 보안(SP)', '영상보안(KT/비대면)', '시스템 보안(SP+KT/비대면)']
    for r in regions: c.execute("INSERT OR IGNORE INTO metadata VALUES ('region', ?)", (r,))
    for cat in categories: c.execute("INSERT OR IGNORE INTO metadata VALUES ('category', ?)", (cat,))
    
    conn.commit()
    conn.close()

init_db()

# 콤마 제거 및 숫자로 변환하는 유틸리티 함수
def clean_num(val):
    if not val: return 0
    return float(str(val).replace(',', ''))

# 2. 통합 웹 애플리케이션 (JS 콤마 처리, 순증 계산, 패스워드 로직 포함)
HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sales Performance Explorer Premium V5</title>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --info: #3b82f6;
            --slate-50: #f8fafc;
            --slate-100: #f1f5f9;
            --slate-200: #e2e8f0;
            --slate-700: #334155;
            --slate-800: #1e293b;
            --glass: rgba(255, 255, 255, 0.7);
            --glass-border: rgba(255, 255, 255, 0.3);
            --shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', 'Malgun Gothic', sans-serif;
            background: linear-gradient(135deg, #f1f5f9 0%, #cbd5e1 100%);
            color: var(--slate-800);
            line-height: 1.6;
            min-height: 100vh;
            padding: 2rem 1rem;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        header {
            text-align: center;
            margin-bottom: 3rem;
        }

        header h1 {
            font-size: 2.5rem;
            font-weight: 800;
            letter-spacing: -0.025em;
            color: var(--slate-800);
            margin-bottom: 0.5rem;
            background: linear-gradient(to right, var(--primary), var(--info));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        header p {
            color: var(--slate-700);
            font-weight: 500;
        }

        /* Tabs */
        .nav-tabs {
            display: flex;
            justify-content: center;
            gap: 1rem;
            margin-bottom: 2rem;
            background: var(--glass);
            backdrop-filter: blur(10px);
            padding: 0.5rem;
            border-radius: 1rem;
            border: 1px solid var(--glass-border);
            box-shadow: var(--shadow);
        }

        .nav-tab {
            padding: 0.75rem 1.5rem;
            border-radius: 0.75rem;
            cursor: pointer;
            font-weight: 600;
            color: var(--slate-700);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            user-select: none;
        }

        .nav-tab:hover { background: rgba(255, 255, 255, 0.5); color: var(--primary); }
        .nav-tab.active { background: var(--primary); color: white; box-shadow: 0 4px 6px -1px rgba(79, 70, 229, 0.4); }

        /* Card System */
        .card {
            background: var(--glass);
            backdrop-filter: blur(16px);
            border: 1px solid var(--glass-border);
            border-radius: 1.5rem;
            padding: 2rem;
            box-shadow: var(--shadow);
            margin-bottom: 2rem;
            display: none;
            animation: slideUp 0.5s ease-out;
        }

        .card.active { display: block; }

        @keyframes slideUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        h2 { font-size: 1.5rem; margin-bottom: 1.5rem; color: var(--slate-800); border-left: 4px solid var(--primary); padding-left: 1rem; }

        /* Form Controls */
        .form-group { margin-bottom: 1.5rem; }
        .label-text { display: block; font-size: 0.875rem; font-weight: 700; margin-bottom: 0.5rem; color: var(--slate-700); }

        .radio-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(130px, 1fr));
            gap: 0.75rem;
            margin-bottom: 2rem;
        }

        .radio-card { position: relative; display: flex; }
        .radio-card input { position: absolute; opacity: 0; cursor: pointer; height: 0; width: 0; }
        .radio-label {
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0.75rem 0.5rem;
            background: white;
            border: 2px solid var(--slate-200);
            border-radius: 0.75rem;
            cursor: pointer;
            font-size: 0.8125rem;
            font-weight: 600;
            transition: all 0.2s;
            text-align: center;
            width: 100%;
            min-height: 60px;
            line-height: 1.3;
            word-break: keep-all;
        }

        .radio-card input:checked + .radio-label { border-color: var(--primary); background: #eef2ff; color: var(--primary); }

        .input-grid {
            display: grid;
            grid-template-columns: repeat(1, 1fr);
            gap: 1.5rem;
            margin-bottom: 2rem;
        }

        .field-container {
            background: rgba(255, 255, 255, 0.5);
            padding: 1.5rem;
            border-radius: 1rem;
            border: 1px solid var(--slate-200);
        }

        /* 7열 그리드로 확장 */
        .field-row {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 0.75rem;
            margin-top: 1rem;
            align-items: end;
        }

        @media (max-width: 1024px) {
            .field-row { grid-template-columns: repeat(4, 1fr); }
        }
        @media (max-width: 640px) {
            .field-row { grid-template-columns: repeat(2, 1fr); }
        }

        input[type="text"], select {
            width: 100%;
            padding: 0.75rem;
            border: 1px solid var(--slate-200);
            border-radius: 0.5rem;
            font-size: 0.875rem;
            transition: border-color 0.2s;
            text-align: right;
        }

        input[type="text"]:focus { outline: none; border-color: var(--primary); ring: 2px solid var(--primary); }
        input.readonly { background: #f1f5f9; cursor: not-allowed; font-weight: 600; }
        input.rate { color: var(--danger); font-weight: 700; background: #fff1f2; }
        input.gap { color: var(--primary); font-weight: 700; background: #eef2ff; }

        .btn {
            width: 100%;
            padding: 1rem;
            border-radius: 0.75rem;
            font-size: 1rem;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.3s;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }

        .btn-primary { background: var(--primary); color: white; }
        .btn-primary:hover { background: var(--primary-hover); transform: translateY(-2px); box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3); }
        .btn-secondary { background: var(--slate-700); color: white; margin-top: 1rem; }
        .btn-secondary:hover { background: var(--slate-800); }

        /* Dashboard specific */
        .dashboard-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.5rem;
        }
        @media (max-width: 900px) { .dashboard-grid { grid-template-columns: 1fr; } }

        .chart-container {
            background: white;
            border-radius: 1rem;
            padding: 1.5rem;
            border: 1px solid var(--slate-200);
            height: 400px;
        }

        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }

        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 1rem;
            border: 1px solid var(--slate-200);
            text-align: center;
        }

        .metric-value { font-size: 1.5rem; font-weight: 800; color: var(--primary); }
        .metric-label { font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--slate-700); margin-top: 0.25rem; }

        /* Excel Upload */
        .upload-area {
            border: 2px dashed var(--slate-200);
            border-radius: 1rem;
            padding: 2rem;
            text-align: center;
            transition: all 0.3s;
            cursor: pointer;
            background: rgba(255, 255, 255, 0.3);
        }
        .upload-area:hover { border-color: var(--primary); background: rgba(79, 70, 229, 0.05); }

        /* Password Modal */
        #password-overlay {
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 23, 42, 0.6); backdrop-filter: blur(8px);
            display: none; justify-content: center; align-items: center; z-index: 1000;
        }
        .pw-card {
            background: white; padding: 2.5rem; border-radius: 1.5rem; width: 90%; max-width: 400px;
            text-align: center; box-shadow: var(--shadow); border: 1px solid var(--glass-border);
        }
        .pw-card h3 { margin-bottom: 1.5rem; color: var(--slate-800); }
        .pw-card input { width: 100%; padding: 1rem; margin-bottom: 1.5rem; border: 2px solid var(--slate-200); border-radius: 0.75rem; font-size: 1.25rem; text-align: center; letter-spacing: 0.5rem; }
        .pw-card .btn-group { display: flex; gap: 0.75rem; }
    </style>
</head>
<body>

<div class="container">
    <header>
        <h1>Performance Explorer <span style="font-size: 0.8rem; vertical-align: middle; background: var(--primary); color: white; padding: 2px 8px; border-radius: 4px;">V5 PRO</span></h1>
        <p>Monthly Sales Forecasting & GAP Analysis System</p>
    </header>

    <div class="nav-tabs">
        <div class="nav-tab active" data-tab="input-tab">📝 실적 입력</div>
        <div class="nav-tab" data-tab="dash-tab">📈 통계 대시보드</div>
        <div class="nav-tab" data-tab="admin-tab">⚙️ 마스터 관리</div>
    </div>

    <!-- 1. Input Side -->
    <div id="input-tab" class="card active">
        <h2>실적 데이터 입력 (GAP 적용)</h2>
        <form id="actualForm">
            <span class="label-text">1. 지역 선택</span>
            <div class="radio-grid" id="regionGroup_user"></div>

            <span class="label-text">2. 카테고리 선택</span>
            <div class="radio-grid" id="categoryGroup_user"></div>

            <div class="input-grid">
                <!-- 신규 실적 -->
                <div class="field-container">
                    <h3 style="font-size: 1rem; color: var(--primary); margin-bottom: 1rem; border-bottom: 1px solid var(--slate-200); padding-bottom: 5px;">🔹 신규 계약 실적 (New Sales)</h3>
                    <div class="field-row">
                        <div style="grid-column: span 1;"><label class="label-text">목표 (자동)</label><input type="text" id="disp_new_target" class="readonly" readonly></div>
                        <div><label class="label-text">4주차 전망</label><input type="text" id="new_actual_4w" oninput="formatAndCalc(this)"></div>
                        <div><label class="label-text">4주차 %</label><input type="text" id="new_rate_4w" class="readonly rate" readonly></div>
                        <div><label class="label-text">마감 전망</label><input type="text" id="new_actual_close" oninput="formatAndCalc(this)"></div>
                        <div><label class="label-text">마감 %</label><input type="text" id="new_rate_close" class="readonly rate" readonly></div>
                        <div><label class="label-text">GAP 금액</label><input type="text" id="new_gap_amt" class="readonly gap" readonly></div>
                        <div><label class="label-text">GAP %</label><input type="text" id="new_gap_rate" class="readonly gap" readonly></div>
                    </div>
                </div>

                <!-- 해지 실적 -->
                <div class="field-container">
                    <h3 style="font-size: 1rem; color: var(--danger); margin-bottom: 1rem; border-bottom: 1px solid var(--slate-200); padding-bottom: 5px;">🔸 해지/이탈 실적 (Cancellations)</h3>
                    <div class="field-row">
                        <div style="grid-column: span 1;"><label class="label-text">목표 (자동)</label><input type="text" id="disp_cancel_target" class="readonly" readonly></div>
                        <div><label class="label-text">4주차 전망</label><input type="text" id="cancel_actual_4w" oninput="formatAndCalc(this)"></div>
                        <div><label class="label-text">4주차 %</label><input type="text" id="cancel_rate_4w" class="readonly rate" readonly></div>
                        <div><label class="label-text">마감 전망</label><input type="text" id="cancel_actual_close" oninput="formatAndCalc(this)"></div>
                        <div><label class="label-text">마감 %</label><input type="text" id="cancel_rate_close" class="readonly rate" readonly></div>
                        <div><label class="label-text">GAP 금액</label><input type="text" id="cancel_gap_amt" class="readonly gap" readonly></div>
                        <div><label class="label-text">GAP %</label><input type="text" id="cancel_gap_rate" class="readonly gap" readonly></div>
                    </div>
                </div>

                <!-- 순증 실적 -->
                <div class="field-container" style="background: #fdf2f2; border-color: var(--primary);">
                    <h3 style="font-size: 1rem; color: var(--primary); margin-bottom: 1rem; border-bottom: 1px solid var(--slate-200); padding-bottom: 5px;">🏆 순증 실적 (Net Performance - 자동계산)</h3>
                    <div class="field-row">
                        <div style="grid-column: span 1;"><label class="label-text">목표 (자동)</label><input type="text" id="disp_net_target" class="readonly" readonly></div>
                        <div><label class="label-text">4주차 전망</label><input type="text" id="net_actual_4w" class="readonly" readonly></div>
                        <div><label class="label-text">4주차 %</label><input type="text" id="net_rate_4w" class="readonly rate" readonly></div>
                        <div><label class="label-text">마감 점망</label><input type="text" id="net_actual_close" class="readonly" readonly></div>
                        <div><label class="label-text">마감 %</label><input type="text" id="net_rate_close" class="readonly rate" readonly></div>
                        <div><label class="label-text">GAP 금액</label><input type="text" id="net_gap_amt" class="readonly gap" readonly></div>
                        <div><label class="label-text">GAP %</label><input type="text" id="net_gap_rate" class="readonly gap" readonly></div>
                    </div>
                </div>
            </div>

            <button type="button" class="btn btn-primary" onclick="submitActuals()">
                🚀 실적 데이터 전송 및 통합하기
            </button>
        </form>
    </div>

    <!-- 2. Dashboard Side -->
    <div id="dash-tab" class="card">
        <h2>전공사 통계 대시보드</h2>
        <div class="form-group">
            <label class="label-text">조회 카테고리</label>
            <select id="dash_category" onchange="loadDashboard()" style="text-align: left; max-width: 300px;">
                <!-- Dynamically filled -->
            </select>
        </div>

        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-value" id="stat_total_target">0</div>
                <div class="metric-label">총 순증 목표</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" id="stat_total_actual">0</div>
                <div class="metric-label">총 마감 전망</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" id="stat_avg_rate">0%</div>
                <div class="metric-label">평균 달성률</div>
            </div>
        </div>

        <div class="dashboard-grid">
            <div class="chart-container">
                <canvas id="performanceChart"></canvas>
            </div>
            <div class="chart-container">
                <canvas id="distributionChart"></canvas>
            </div>
        </div>

        <button type="button" class="btn btn-secondary" onclick="exportData()">
            📥 양식 동기화 엑셀(XLSX) 다운로드
        </button>
    </div>

    <!-- 3. Admin Side -->
    <div id="admin-tab" class="card">
        <h2>마스터 및 목표 관리</h2>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem;">
            <div>
                <h3 style="font-size: 1rem; margin-bottom: 1rem;">📍 개별 목표 설정</h3>
                <form id="targetForm">
                    <span class="label-text">지역 선택</span>
                    <div class="radio-grid" id="regionGroup_admin"></div>
                    <span class="label-text">카테고리 선택</span>
                    <div class="radio-grid" id="categoryGroup_admin"></div>
                    
                    <div class="field-row" style="margin-bottom: 1.5rem; grid-template-columns: repeat(3, 1fr);">
                        <div><label class="label-text">신규 목표</label><input type="text" id="admin_new_target" oninput="formatAdmin(this)"></div>
                        <div><label class="label-text">해지 목표</label><input type="text" id="admin_cancel_target" oninput="formatAdmin(this)"></div>
                        <div><label class="label-text">순증 목표</label><input type="text" id="admin_net_target" class="readonly" readonly></div>
                    </div>
                    <button type="button" class="btn btn-primary" onclick="submitTargets()">목표 저장</button>
                </form>
            </div>

            <div>
                <h3 style="font-size: 1rem; margin-bottom: 1rem;">📊 엑셀 대량 업로드</h3>
                <div class="upload-area" onclick="triggerUpload()">
                    <div style="font-size: 2rem; margin-bottom: 0.5rem;">📁</div>
                    <p style="font-weight: 600;">클릭하여 엑셀 파일 선택</p>
                    <p style="font-size: 0.75rem; color: var(--slate-700); margin-top: 0.5rem;">목표 데이터(targets) 또는 실적 데이터(actuals) 벌크 업데이트</p>
                </div>
                <input type="file" id="excelFile" style="display: none;" onchange="handleFileUpload(this)">
            </div>
        </div>
    </div>

    <!-- Password Modal -->
    <div id="password-overlay">
        <div class="pw-card">
            <h3>🔒 관리자 인증</h3>
            <p style="font-size: 0.875rem; color: var(--slate-700); margin-bottom: 1rem;">비밀번호를 입력하세요.</p>
            <input type="password" id="admin_pw" placeholder="****" onkeydown="if(event.key==='Enter') verifyPw()">
            <div class="btn-group">
                <button type="button" class="btn btn-secondary" onclick="closePwModal()" style="margin-top:0;">취소</button>
                <button type="button" class="btn btn-primary" onclick="verifyPw()" style="margin-top:0;">확인</button>
            </div>
        </div>
    </div>
</div>

<script>
    let regions = [];
    let categories = [];
    let mainChart = null;
    let mainPie = null;

    window.onload = () => {
        fetch('/api/metadata')
            .then(res => res.json())
            .then(data => {
                regions = data.regions;
                categories = data.categories;
                initUI();
                setupEventListeners();
                fetchTarget();
            });
    };

    function initUI() {
        renderRadios('regionGroup_user', 'region', regions, 'fetchTarget()');
        renderRadios('categoryGroup_user', 'category', categories, 'fetchTarget()');
        renderRadios('regionGroup_admin', 'region_admin', regions, '');
        renderRadios('categoryGroup_admin', 'category_admin', categories, '');
        
        const dashCat = document.getElementById('dash_category');
        dashCat.innerHTML = categories.map(c => `<option value="${c}">${c}</option>`).join('');
    }

    function renderRadios(containerId, name, items, onchange) {
        const container = document.getElementById(containerId);
        container.innerHTML = items.map((item, idx) => `
            <div class="radio-card">
                <input type="radio" id="${containerId}_${idx}" name="${name}" value="${item}" ${idx===0?'checked':''} onchange="${onchange}">
                <label class="radio-label" for="${containerId}_${idx}">${item}</label>
            </div>
        `).join('');
    }

    function setupEventListeners() {
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.addEventListener('click', function(e) {
                const targetId = this.getAttribute('data-tab');
                if(targetId === 'admin-tab') {
                    showPwModal();
                } else {
                    switchTab(targetId, this);
                }
            });
        });
    }

    function showPwModal() {
        document.getElementById('password-overlay').style.display = 'flex';
        document.getElementById('admin_pw').value = '';
        document.getElementById('admin_pw').focus();
    }

    function closePwModal() {
        document.getElementById('password-overlay').style.display = 'none';
    }

    function verifyPw() {
        const pw = document.getElementById('admin_pw').value;
        if(pw === "1234") {
            closePwModal();
            const tab = document.querySelector('.nav-tab[data-tab="admin-tab"]');
            switchTab('admin-tab', tab);
        } else {
            alert("비밀번호가 틀렸습니다.");
            document.getElementById('admin_pw').value = '';
            document.getElementById('admin_pw').focus();
        }
    }

    function switchTab(targetId, tabEl) {
        document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.card').forEach(c => c.classList.remove('active'));
        
        tabEl.classList.add('active');
        const targetCard = document.getElementById(targetId);
        if (targetCard) targetCard.classList.add('active');

        if(targetId === 'dash-tab') loadDashboard();
    }

    function fmt(n) { return n ? n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") : "0"; }
    function p(v) { return v ? v.toString().replace(/,/g, '') : "0"; }
    function getVal(id) { return parseFloat(p(document.getElementById(id).value)) || 0; }

    // V5 핵심: GAP 분석 기능 통합
    function calcRateAndGap(targetId, actualId, rateId, gapAmtId, gapRateId) {
        let target = getVal(targetId);
        let actual = getVal(actualId);
        
        if(target !== 0) {
            document.getElementById(rateId).value = ((actual / target) * 100).toFixed(1) + "%";
            if(gapAmtId) {
                let gapAmt = actual - target;
                document.getElementById(gapAmtId).value = (gapAmt >= 0 ? "+" : "") + fmt(gapAmt);
                document.getElementById(gapRateId).value = ((gapAmt / target) * 100).toFixed(1) + "%";
            }
        } else {
            document.getElementById(rateId).value = "0%";
            if(gapAmtId) { document.getElementById(gapAmtId).value = "0"; document.getElementById(gapRateId).value = "0%"; }
        }
    }

    function formatAndCalc(el) {
        const val = el.value.replace(/[^0-9-]/g, '');
        if(val) el.value = fmt(val);
        
        // 순증 계산
        let new4w = getVal('new_actual_4w'); let can4w = getVal('cancel_actual_4w');
        document.getElementById('net_actual_4w').value = fmt(new4w - can4w);

        let newClose = getVal('new_actual_close'); let canClose = getVal('cancel_actual_close');
        document.getElementById('net_actual_close').value = fmt(newClose - canClose);

        // 신규 GAP
        calcRateAndGap('disp_new_target', 'new_actual_4w', 'new_rate_4w', null, null);
        calcRateAndGap('disp_new_target', 'new_actual_close', 'new_rate_close', 'new_gap_amt', 'new_gap_rate');
        
        // 해지 GAP
        calcRateAndGap('disp_cancel_target', 'cancel_actual_4w', 'cancel_rate_4w', null, null);
        calcRateAndGap('disp_cancel_target', 'cancel_actual_close', 'cancel_rate_close', 'cancel_gap_amt', 'cancel_gap_rate');
        
        // 순증 GAP
        calcRateAndGap('disp_net_target', 'net_actual_4w', 'net_rate_4w', null, null);
        calcRateAndGap('disp_net_target', 'net_actual_close', 'net_rate_close', 'net_gap_amt', 'net_gap_rate');
    }

    function formatAdmin(el) {
        const val = el.value.replace(/[^0-9-]/g, '');
        if(val) el.value = fmt(val);
        document.getElementById('admin_net_target').value = fmt(getVal('admin_new_target') - getVal('admin_cancel_target'));
    }

    function fetchTarget() {
        const r = document.querySelector('input[name="region"]:checked').value;
        const c = document.querySelector('input[name="category"]:checked').value;
        fetch(`/api/get_target?region=${r}&category=${c}`)
            .then(res => res.json())
            .then(data => {
                document.getElementById('disp_new_target').value = fmt(data.new_target);
                document.getElementById('disp_cancel_target').value = fmt(data.cancel_target);
                document.getElementById('disp_net_target').value = fmt(getVal('disp_new_target') - getVal('disp_cancel_target'));
                formatAndCalc(document.getElementById('new_actual_4w'));
            });
    }

    function submitActuals() {
        const fd = new FormData();
        fd.append('region', document.querySelector('input[name="region"]:checked').value);
        fd.append('category', document.querySelector('input[name="category"]:checked').value);
        fd.append('new_actual_4w', getVal('new_actual_4w'));
        fd.append('new_actual_close', getVal('new_actual_close'));
        fd.append('cancel_actual_4w', getVal('cancel_actual_4w'));
        fd.append('cancel_actual_close', getVal('cancel_actual_close'));

        fetch('/submit_actual', { method: 'POST', body: fd })
            .then(res => res.text()).then(m => { alert(m); location.reload(); });
    }

    function submitTargets() {
        const fd = new FormData();
        fd.append('region', document.querySelector('input[name="region_admin"]:checked').value);
        fd.append('category', document.querySelector('input[name="category_admin"]:checked').value);
        fd.append('new_target', getVal('admin_new_target'));
        fd.append('cancel_target', getVal('admin_cancel_target'));

        fetch('/submit_target', { method: 'POST', body: fd })
            .then(res => res.text()).then(m => { alert(m); fetchTarget(); });
    }

    function loadDashboard() {
        const cat = document.getElementById('dash_category').value;
        fetch(`/api/dashboard?category=${cat}`)
            .then(res => res.json())
            .then(data => {
                const labels = data.map(d => d.region);
                const targets = data.map(d => d.net_target);
                const actuals = data.map(d => d.net_actual_close || 0);
                
                let sumT = targets.reduce((a,b)=>a+b, 0);
                let sumA = actuals.reduce((a,b)=>a+b, 0);
                
                document.getElementById('stat_total_target').innerText = fmt(sumT);
                document.getElementById('stat_total_actual').innerText = fmt(sumA);
                document.getElementById('stat_avg_rate').innerText = (sumT ? (sumA/sumT*100).toFixed(1) : 0) + "%";

                renderCharts(labels, targets, actuals);
            });
    }

    function renderCharts(labels, targets, actuals) {
        if(mainChart) mainChart.destroy();
        if(mainPie) mainPie.destroy();

        const ctx = document.getElementById('performanceChart').getContext('2d');
        mainChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    { label: '목표 (Target)', data: targets, backgroundColor: 'rgba(79, 70, 229, 0.2)', borderColor: '#4f46e5', borderWidth: 2 },
                    { label: '전망 (Projection)', data: actuals, backgroundColor: 'rgba(16, 185, 129, 0.6)', borderRadius: 4 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { font: { family: 'Inter', weight: '600' } } } },
                scales: { y: { beginAtZero: true, grid: { color: '#f1f5f9' } } }
            }
        });

        const ctxPie = document.getElementById('distributionChart').getContext('2d');
        mainPie = new Chart(ctxPie, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: actuals,
                    backgroundColor: ['#4f46e5', '#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { 
                    legend: { position: 'right', labels: { boxWidth: 12, padding: 15, font: { weight: '600' } } },
                }
            }
        });
    }

    function exportData() { location.href = '/download'; }

    function triggerUpload() { document.getElementById('excelFile').click(); }
    function handleFileUpload(input) {
        if(!input.files.length) return;
        const file = input.files[0];
        const fd = new FormData();
        fd.append('file', file);
        
        const type = confirm("목표 데이터(targets) 업로드입니까? (취소 시 실적actuals 업로드)") ? 'target' : 'actual';
        fd.append('type', type);

        fetch('/api/upload_excel', { method: 'POST', body: fd })
            .then(res => res.json())
            .then(data => { alert(data.msg); location.reload(); })
            .catch(() => alert("업로드 실패"));
    }
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/api/get_target', methods=['GET'])
def get_target():
    region = request.args.get('region')
    category = request.args.get('category')
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT new_target, cancel_target FROM targets WHERE region=? AND category=?", (region, category)).fetchone()
    conn.close()
    if row: return jsonify(dict(row))
    return jsonify({"new_target": 0, "cancel_target": 0})

@app.route('/api/dashboard', methods=['GET'])
def get_dashboard():
    category = request.args.get('category')
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    # 대시보드 시각화를 '순증(신규-해지)' 기준으로 처리
    query = """
    SELECT t.region, (IFNULL(t.new_target, 0) - IFNULL(t.cancel_target, 0)) as net_target, 
           (SELECT (IFNULL(new_actual_close, 0) - IFNULL(cancel_actual_close, 0)) FROM actuals a WHERE a.region = t.region AND a.category = t.category ORDER BY timestamp DESC LIMIT 1) as net_actual_close
    FROM targets t WHERE t.category=?
    """
    rows = conn.execute(query, (category,)).fetchall()
    conn.close()
    
    # 정의된 REGIONS_ORDER 순서대로 데이터 정렬
    results = [dict(row) for row in rows]
    sorted_results = sorted(results, key=lambda x: REGIONS_ORDER.index(x['region']) if x['region'] in REGIONS_ORDER else 999)
    
    return jsonify(sorted_results)

@app.route('/submit_target', methods=['POST'])
def submit_target():
    data = (request.form.get('region'), request.form.get('category'), clean_num(request.form.get('new_target')), clean_num(request.form.get('cancel_target')))
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT OR REPLACE INTO targets (region, category, new_target, cancel_target) VALUES (?, ?, ?, ?)", data)
    conn.commit()
    conn.close()
    return f"[{data[0]}] {data[1]} 목표가 설정되었습니다."

@app.route('/api/metadata', methods=['GET'])
def get_metadata():
    # 정의된 순서대로 반환하여 UI 일관성 유지
    return jsonify({"regions": REGIONS_ORDER, "categories": CATEGORIES_ORDER})

@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files: return jsonify({"msg": "파일이 없습니다."}), 400
    file = request.files['file']
    upload_type = request.form.get('type') # 'target' or 'actual'
    
    try:
        df = pd.read_excel(file)
        conn = sqlite3.connect(DB_NAME)
        if upload_type == 'target':
            for _, row in df.iterrows():
                conn.execute("INSERT OR REPLACE INTO targets (region, category, new_target, cancel_target) VALUES (?, ?, ?, ?)", 
                             (row['지역'], row['카테고리'], row['신규목표'], row['해지목표']))
        else:
            for _, row in df.iterrows():
                conn.execute("INSERT INTO actuals (region, category, new_actual_4w, new_actual_close, cancel_actual_4w, cancel_actual_close) VALUES (?, ?, ?, ?, ?, ?)", 
                             (row['지역'], row['카테고리'], row['신규4주차'], row['신규마감'], row['해지4주차'], row['해지마감']))
        conn.commit()
        conn.close()
        return jsonify({"msg": f"성공적으로 {len(df)}건의 데이터를 업로드했습니다."})
    except Exception as e:
        return jsonify({"msg": f"오류 발생: {str(e)}"}), 500

@app.route('/submit_actual', methods=['POST'])
def submit_actual():
    data = (request.form.get('region'), request.form.get('category'),
            clean_num(request.form.get('new_actual_4w')), clean_num(request.form.get('new_actual_close')),
            clean_num(request.form.get('cancel_actual_4w')), clean_num(request.form.get('cancel_actual_close')))
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT INTO actuals (region, category, new_actual_4w, new_actual_close, cancel_actual_4w, cancel_actual_close) VALUES (?, ?, ?, ?, ?, ?)", data)
    conn.commit()
    conn.close()
    return f"[{data[0]}] {data[1]} 실적이 저장되었습니다."

# 달성률 포맷팅용 헬퍼 함수
def calc_rate(actual, target):
    if target == 0 or pd.isna(target): return '0%'
    return f"{(actual / target * 100):.1f}%"

def calc_gap_rate(actual, target):
    if target == 0 or pd.isna(target): return '0%'
    return f"{((actual - target) / target * 100):.1f}%"

@app.route('/download')
def download():
    conn = sqlite3.connect(DB_NAME)
    query = """
    SELECT a.region, a.category, 
           IFNULL(t.new_target, 0) as new_target, a.new_actual_4w, a.new_actual_close,
           IFNULL(t.cancel_target, 0) as cancel_target, a.cancel_actual_4w, a.cancel_actual_close,
           a.timestamp
    FROM actuals a
    LEFT JOIN targets t ON a.region = t.region AND a.category = t.category
    ORDER BY a.timestamp DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty: return "아직 데이터가 없습니다."

    # 순증 계산
    df['net_target'] = df['new_target'] - df['cancel_target']
    df['net_actual_4w'] = df['new_actual_4w'] - df['cancel_actual_4w']
    df['net_actual_close'] = df['new_actual_close'] - df['cancel_actual_close']

    # ★ 핵심: 엑셀 파일 헤더 병합을 위한 Pandas MultiIndex 구조 세팅
    columns = pd.MultiIndex.from_tuples([
        ('기본정보', '지역'), ('기본정보', '카테고리'),
        ('신규', '목표'), ('신규', '4주차 실적'), ('신규', '4주차 달성률'), ('신규', '마감 실적'), ('신규', '마감 달성률'), ('신규', 'GAP 금액'), ('신규', 'GAP %'),
        ('해지', '목표'), ('해지', '4주차 실적'), ('해지', '4주차 달성률'), ('해지', '마감 실적'), ('해지', '마감 달성률'), ('해지', 'GAP 금액'), ('해지', 'GAP %'),
        ('순증', '목표'), ('순증', '4주차 실적'), ('순증', '4주차 달성률'), ('순증', '마감 실적'), ('순증', '마감 달성률'), ('순증', 'GAP 금액'), ('순증', 'GAP %'),
        ('시스템', '입력시간')
    ])
    
    export_df = pd.DataFrame(columns=columns)
    
    export_df[('기본정보', '지역')] = df['region']
    export_df[('기본정보', '카테고리')] = df['category']
    
    # 신규
    export_df[('신규', '목표')] = df['new_target']
    export_df[('신규', '4주차 실적')] = df['new_actual_4w']
    export_df[('신규', '4주차 달성률')] = df.apply(lambda r: calc_rate(r['new_actual_4w'], r['new_target']), axis=1)
    export_df[('신규', '마감 실적')] = df['new_actual_close']
    export_df[('신규', '마감 달성률')] = df.apply(lambda r: calc_rate(r['new_actual_close'], r['new_target']), axis=1)
    export_df[('신규', 'GAP 금액')] = df['new_actual_close'] - df['new_target']
    export_df[('신규', 'GAP %')] = df.apply(lambda r: calc_gap_rate(r['new_actual_close'], r['new_target']), axis=1)

    # 해지
    export_df[('해지', '목표')] = df['cancel_target']
    export_df[('해지', '4주차 실적')] = df['cancel_actual_4w']
    export_df[('해지', '4주차 달성률')] = df.apply(lambda r: calc_rate(r['cancel_actual_4w'], r['cancel_target']), axis=1)
    export_df[('해지', '마감 실적')] = df['cancel_actual_close']
    export_df[('해지', '마감 달성률')] = df.apply(lambda r: calc_rate(r['cancel_actual_close'], r['cancel_target']), axis=1)
    export_df[('해지', 'GAP 금액')] = df['cancel_actual_close'] - df['cancel_target']
    export_df[('해지', 'GAP %')] = df.apply(lambda r: calc_gap_rate(r['cancel_actual_close'], r['cancel_target']), axis=1)

    # 순증
    export_df[('순증', '목표')] = df['net_target']
    export_df[('순증', '4주차 실적')] = df['net_actual_4w']
    export_df[('순증', '4주차 달성률')] = df.apply(lambda r: calc_rate(r['net_actual_4w'], r['net_target']), axis=1)
    export_df[('순증', '마감 실적')] = df['net_actual_close']
    export_df[('순증', '마감 달성률')] = df.apply(lambda r: calc_rate(r['net_actual_close'], r['net_target']), axis=1)
    export_df[('순증', 'GAP 금액')] = df['net_actual_close'] - df['net_target']
    export_df[('순증', 'GAP %')] = df.apply(lambda r: calc_gap_rate(r['net_actual_close'], r['net_target']), axis=1)
    
    export_df[('시스템', '입력시간')] = df['timestamp']

    # 메모리에 엑셀 생성
    excel_file = io.BytesIO()
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='마감회의자료_취합')
    excel_file.seek(0)

    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('회의자료_동기화결과.xlsx', excel_file.getvalue())
    memory_file.seek(0)

    filename = f"{datetime.now().strftime('%Y%m%d_%H%M')}_마감취합_V5.zip"
    return send_file(memory_file, download_name=filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
