# 見積書作成アプリ
import streamlit as st
import pandas as pd
import datetime
import os
import json
import openpyxl
import traceback
from estimate_excel_writer import write_estimate_to_excel

# ページ設定
st.set_page_config(page_title="見積書作成アプリ", layout="wide")
st.title("見積書作成アプリ")

# 定数定義
ISSUER_LIST = ["須藤 竜平", "本間 清昭", "片岡 啓明", "青山 泰", "中角 明子"]
EXCEL_FILENAME = "見積管理データ.xlsx"
DATA_FOLDER = "data"

# セッション状態の初期化
def init_session_state():
    """セッション状態を初期化"""
    defaults = {
        "明細リスト": [],
        "編集対象": None,
        "発行日": None,
        "見積No": "",
        "案件名": "",
        "選択された顧客会社名": "",
        "選択された顧客担当者": "",
        "選択された顧客部署名": "",
        "売上額自動更新": True,
        "アクティブタブ": "① 案件一覧",
        "発行者名": ISSUER_LIST[0],
        "備考": "",
        "メモ": "",
        "状況": "見積中",
        "受注日": None,
        "納品日": None,
        "売上額": 0,
        "仕入額": 0,
        "粗利": 0,
        "粗利率": 0,
        "係数機能使用": False  # 係数機能のフラグを追加
    }
    
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default

# データ読み込み関数
@st.cache_resource
def load_data():
    """JSONファイルからデータを読み込む"""
    try:
        # 顧客データの読み込み
        顧客一覧 = load_customers_json()
        顧客一覧_df = pd.DataFrame(顧客一覧) if 顧客一覧 else pd.DataFrame(columns=["顧客No", "顧客会社名", "顧客部署名", "顧客担当者", "顧客住所"])
        
        # 案件データの読み込み
        案件一覧 = load_all_projects()
        案件一覧_df = pd.DataFrame(案件一覧) if 案件一覧 else pd.DataFrame(columns=["見積No", "案件名", "顧客会社名", "顧客部署名", "顧客担当者", "発行日", "受注日", "納品日", "売上額", "仕入額", "粗利", "粗利率", "状況", "発行者名", "メモ"])
        
        # 商品データの読み込み
        品名一覧 = load_products_json()
        品名一覧_df = pd.DataFrame(品名一覧) if 品名一覧 else pd.DataFrame(columns=["品名", "単位", "単価", "備考"])
        
        return 顧客一覧_df, 案件一覧_df, 品名一覧_df
        
    except Exception as e:
        st.error(f"データ読み込みエラー: {e}")
        return create_empty_dataframes()

def create_empty_dataframes():
    """空のデータフレームを作成"""
    顧客一覧 = pd.DataFrame(columns=["顧客No", "顧客会社名", "顧客部署名", "顧客担当者", "顧客住所"])
    案件一覧 = pd.DataFrame(columns=["見積No", "案件名", "顧客会社名", "顧客部署名", "顧客担当者", "発行日", "受注日", "納品日", "売上額", "仕入額", "粗利", "粗利率", "状況", "発行者名", "メモ"])
    品名一覧 = pd.DataFrame(columns=["品名", "単位", "単価", "備考"])
    return 顧客一覧, 案件一覧, 品名一覧

# JSON関連の関数
def save_meisai_as_json(見積No, data):
    """明細データをJSONファイルに保存（数値フィールド修正版）"""
    try:
        os.makedirs(DATA_FOLDER, exist_ok=True)
        ファイルパス = os.path.join(DATA_FOLDER, f"{見積No}.json")
        
        # 顧客担当者名のスペースを正規化（半角スペースに統一）
        顧客担当者 = st.session_state.get("選択された顧客担当者", "")
        正規化顧客担当者 = 顧客担当者.replace("　", " ").strip()
        
        # 顧客住所を取得（顧客一覧から検索）
        顧客住所 = ""
        try:
            顧客一覧, _, _ = load_excel_data(EXCEL_FILENAME)
            顧客会社名 = st.session_state.get("選択された顧客会社名", "")
            if 顧客会社名 and not 顧客一覧.empty:
                該当顧客 = 顧客一覧[顧客一覧["顧客会社名"] == 顧客会社名]
                if not 該当顧客.empty:
                    顧客住所 = 該当顧客.iloc[0].get("顧客住所", "")
        except:
            pass
        
        # 明細から合計金額を計算（分類項目除外・数値チェック強化）
        明細リスト = data.get("明細リスト", st.session_state.get("明細リスト", []))
        明細合計 = 0
        
        # 明細リストの数値フィールドを正規化
        正規化明細リスト = []
        for item in 明細リスト:
            正規化item = item.copy()
            
            # 分類項目の場合は数値フィールドを0に統一
            if item.get("分類", False):
                正規化item["数量"] = 0
                正規化item["単価"] = 0
                正規化item["金額"] = 0
            else:
                # 商品項目の場合は数値として確実に保存
                try:
                    正規化item["数量"] = int(item.get("数量", 0)) if item.get("数量") not in ["", None] else 0
                except (ValueError, TypeError):
                    正規化item["数量"] = 0
                
                try:
                    正規化item["単価"] = float(item.get("単価", 0)) if item.get("単価") not in ["", None] else 0
                except (ValueError, TypeError):
                    正規化item["単価"] = 0
                
                try:
                    正規化item["金額"] = float(item.get("金額", 0)) if item.get("金額") not in ["", None] else 0
                except (ValueError, TypeError):
                    正規化item["金額"] = 0
                
                # 明細合計に加算（分類項目は除外）
                明細合計 += 正規化item["金額"]
            
            # その他のフィールドも文字列として確実に保存
            正規化item["品名"] = str(item.get("品名", ""))
            正規化item["単位"] = str(item.get("単位", ""))
            正規化item["備考"] = str(item.get("備考", ""))
            正規化item["売上先部署"] = str(item.get("売上先部署", ""))
            正規化item["分類"] = bool(item.get("分類", False))
            
            正規化明細リスト.append(正規化item)
        
        # 売上額自動更新が有効な場合は明細合計を使用
        if st.session_state.get("売上額自動更新", True):
            売上額 = 明細合計
        else:
            売上額 = st.session_state.get("売上額", 明細合計)
        
        # 粗利の再計算
        仕入額 = int(st.session_state.get("仕入額", 0))
        粗利 = 売上額 - 仕入額
        粗利率 = (粗利 / 売上額 * 100) if 売上額 > 0 else 0
        
        # セッション状態も更新
        st.session_state["売上額"] = 売上額
        st.session_state["粗利"] = 粗利
        st.session_state["粗利率"] = 粗利率
        
        # 見積データ全体を保存（数値フィールド正規化済み）
        保存データ = {
            "見積No": str(st.session_state.get("見積No", "")),
            "案件名": str(st.session_state.get("案件名", "")),
            "発行日": str(st.session_state.get("発行日", datetime.date.today())),
            "顧客会社名": str(st.session_state.get("選択された顧客会社名", "")),
            "顧客部署名": str(st.session_state.get("選択された顧客部署名", "")),
            "顧客担当者": 正規化顧客担当者,
            "顧客住所": str(顧客住所),
            "発行者名": str(st.session_state.get("発行者名", "")),
            "担当部署": str(st.session_state.get("担当部署", "")),
            "備考": str(data.get("備考", st.session_state.get("備考", ""))),
            "明細リスト": 正規化明細リスト,  # 正規化済みの明細リスト
            "状況": str(st.session_state.get("状況", "見積中")),
            "受注日": str(st.session_state.get("受注日", "") or ""),
            "納品日": str(st.session_state.get("納品日", "") or ""),
            "売上額": int(売上額),
            "仕入額": int(仕入額),
            "粗利": int(粗利),
            "粗利率": float(粗利率),
            "メモ": str(st.session_state.get("メモ", ""))
        }
        
        # 上書き処理があれば先に実行
        上書き処理 = st.session_state.get("上書き処理")
        if 上書き処理 and 上書き処理.get("実行予定"):
            旧見積No = 上書き処理["旧見積No"]
            新見積No = 上書き処理["新見積No"]
            
            if 旧見積No != 新見積No:
                # 旧ファイルを削除
                旧ファイルパス = os.path.join(DATA_FOLDER, f"{旧見積No}.json")
                try:
                    if os.path.exists(旧ファイルパス):
                        os.remove(旧ファイルパス)
                        st.success(f"旧データ（{旧見積No}）を削除しました")
                    
                    # 上書き処理をクリア
                    del st.session_state["上書き処理"]
                    
                except Exception as e:
                    st.error(f"旧ファイル削除エラー: {e}")
        
        # 新しいファイルを保存
        with open(ファイルパス, "w", encoding="utf-8") as f:
            json.dump(保存データ, f, ensure_ascii=False, indent=2, default=str)  # default=strを追加
        
        return True
        
    except Exception as e:
        st.error(f"JSONファイル保存エラー: {e}")
        return False

def safe_date(data, key):
    """日付データを安全に変換する"""
    try:
        return pd.to_datetime(data[key]).date() if key in data and data[key] else None
    except:
        return None

def auto_load_json_by_estimate_no(見積No, auto_rerun=True): 
    """見積番号に対応するJSONファイルを自動読み込み（数値変換強化版）"""
    try:
        ファイルパス = os.path.join(DATA_FOLDER, f"{見積No}.json")
        
        if not os.path.exists(ファイルパス):
            st.error(f"JSONファイルが見つかりません: {ファイルパス}")
            return False

        with open(ファイルパス, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        if not isinstance(data, dict):
            st.error("JSONファイルのデータ形式が正しくありません")
            return False

        # 明細リストの数値フィールドを修正
        明細リスト = data.get("明細リスト", [])
        修正済み明細リスト = []
        
        for item in 明細リスト:
            修正済みitem = item.copy()
            
            # 分類項目の場合は数値フィールドを0に統一
            if item.get("分類", False):
                修正済みitem["数量"] = 0
                修正済みitem["単価"] = 0
                修正済みitem["金額"] = 0
            else:
                # 商品項目の場合は安全に数値変換
                try:
                    数量 = item.get("数量", 0)
                    if 数量 == "" or 数量 is None:
                        修正済みitem["数量"] = 0
                    else:
                        修正済みitem["数量"] = int(数量)
                except (ValueError, TypeError):
                    修正済みitem["数量"] = 0
                
                try:
                    単価 = item.get("単価", 0)
                    if 単価 == "" or 単価 is None:
                        修正済みitem["単価"] = 0
                    else:
                        修正済みitem["単価"] = float(単価)
                except (ValueError, TypeError):
                    修正済みitem["単価"] = 0
                
                try:
                    金額 = item.get("金額", 0)
                    if 金額 == "" or 金額 is None:
                        修正済みitem["金額"] = 0
                    else:
                        修正済みitem["金額"] = float(金額)
                except (ValueError, TypeError):
                    修正済みitem["金額"] = 0
            
            # その他のフィールドも安全に変換
            修正済みitem["品名"] = str(item.get("品名", ""))
            修正済みitem["単位"] = str(item.get("単位", ""))
            修正済みitem["備考"] = str(item.get("備考", ""))
            修正済みitem["売上先部署"] = str(item.get("売上先部署", ""))
            修正済みitem["分類"] = bool(item.get("分類", False))
            
            修正済み明細リスト.append(修正済みitem)
        
        # 修正済み明細リストをdataに反映
        data["明細リスト"] = 修正済み明細リスト

        # データの検証と正規化
        valid_status = ["見積中", "受注", "納品済", "請求済", "不採用", "失注"]
        status = data.get("状況", "見積中")
        if status not in valid_status:
            status = "見積中"

        valid_issuer = ISSUER_LIST
        issuer = data.get("発行者名", ISSUER_LIST[0])
        if issuer not in valid_issuer:
            issuer = ISSUER_LIST[0]

        # 必須情報のセッション登録
        st.session_state["見積No"] = data.get("見積No", "")
        st.session_state["案件名"] = data.get("案件名", "")
        st.session_state["発行日"] = safe_date(data, "発行日") or datetime.date.today()
        st.session_state["選択された顧客会社名"] = data.get("顧客会社名", "")
        st.session_state["選択された顧客部署名"] = data.get("顧客部署名", "")
        st.session_state["選択された顧客担当者"] = data.get("顧客担当者", "")
        st.session_state["選択された顧客住所"] = data.get("顧客住所", "")
        st.session_state["発行者名"] = issuer
        st.session_state["担当部署"] = data.get("担当部署", "")
        st.session_state["明細リスト"] = 修正済み明細リスト  # 修正済みの明細リストを使用
        st.session_state["備考"] = data.get("備考", "")
        st.session_state["売上額"] = int(data.get("売上額", 0))
        st.session_state["仕入額"] = int(data.get("仕入額", 0))
        st.session_state["粗利"] = int(data.get("粗利", 0))
        st.session_state["粗利率"] = float(data.get("粗利率", 0.0))
        st.session_state["状況"] = status
        st.session_state["受注日"] = safe_date(data, "受注日")
        st.session_state["納品日"] = safe_date(data, "納品日")
        st.session_state["メモ"] = data.get("メモ", "")
        
        # auto_rerunパラメータで再実行を制御
        if auto_rerun:
            st.rerun()
        return True
        
    except FileNotFoundError:
        st.error(f"ファイルが見つかりません: {見積No}.json")
        return False
    except json.JSONDecodeError as e:
        st.error(f"JSONファイルの読み込みエラー: {e}")
        return False
    except Exception as e:
        st.error(f"予期しないエラーが発生しました: {e}")
        return False

# Excel保存関数（改善版）
def save_to_excel(data, sheet_name, filename=EXCEL_FILENAME):
    """データをExcelファイルに保存（改善版）"""
    try:
        if os.path.exists(filename):
            # 既存のファイルを読み込む
            with pd.ExcelFile(filename, engine='openpyxl') as xls:
                # 各シートを辞書として保存
                sheets = {}
                for name in xls.sheet_names:
                    if name != sheet_name:  # 更新対象以外のシートを保存
                        sheets[name] = xls.parse(name)
                
                # 更新対象のシートを追加
                sheets[sheet_name] = data
            
            # すべてのシートを書き戻す
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                for name, df in sheets.items():
                    df.to_excel(writer, sheet_name=name, index=False)
        else:
            # 新規ファイルの場合
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                data.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"Excel保存エラー: {e}")
        return False

# ユーティリティ関数
def count_same_date_projects_from_json(発行日_str):
    """JSONファイルから同日案件数をカウント（修正版）"""
    count = 0
    if os.path.exists(DATA_FOLDER):
        try:
            json_files = [f for f in os.listdir(DATA_FOLDER) if f.endswith('.json')]
            for file in json_files:
                try:
                    ファイルパス = os.path.join(DATA_FOLDER, file)
                    with open(ファイルパス, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    
                    if isinstance(data, dict):
                        # ファイル内の発行日を取得
                        file_発行日 = data.get("発行日", "")
                        
                        # 日付の正規化（異なる形式に対応）
                        if isinstance(file_発行日, str):
                            # "2025-03-15" 形式の場合
                            if len(file_発行日) == 10 and "-" in file_発行日:
                                try:
                                    file_date = datetime.datetime.strptime(file_発行日, "%Y-%m-%d").date()
                                    file_発行日_str = file_date.strftime('%Y%m%d')
                                except:
                                    continue
                            # "20250315" 形式の場合
                            elif len(file_発行日) == 8 and file_発行日.isdigit():
                                file_発行日_str = file_発行日
                            else:
                                continue
                        else:
                            continue
                        
                        # 同じ発行日かチェック
                        if file_発行日_str == 発行日_str:
                            count += 1
                            
                except Exception as e:
                    # 個別ファイルのエラーは無視して継続
                    continue
        except Exception as e:
            # ディレクトリアクセスエラーも無視
            pass
    
    return count

def generate_estimate_no(発行日):
    """見積番号を生成（None対応版）"""
    # 発行日がNoneの場合は空文字列を返す
    if 発行日 is None:
        return ""
    
    発行日_str = 発行日.strftime('%Y%m%d')
    
    # 同日の既存案件数を正確にカウント
    同日案件数 = count_same_date_projects_from_json(発行日_str)
    
    # 連番を生成（既存案件数 + 1）
    連番 = 同日案件数 + 1
    
    # 見積番号を生成
    見積No = f"{発行日_str}{str(連番).zfill(3)}"
    
    return 見積No

def check_estimate_no_exists(見積No):
    """見積番号が既に存在するかチェック"""
    ファイルパス = os.path.join(DATA_FOLDER, f"{見積No}.json")
    return os.path.exists(ファイルパス)

def generate_unique_estimate_no(発行日):
    """重複しない見積番号を生成（None対応版）"""
    # 発行日がNoneの場合は空文字列を返す
    if 発行日 is None:
        return ""
    
    発行日_str = 発行日.strftime('%Y%m%d')
    
    # 連番を1から開始して、重複しない番号を見つける
    連番 = 1
    while True:
        見積No = f"{発行日_str}{str(連番).zfill(3)}"
        
        # この番号が既に存在するかチェック
        if not check_estimate_no_exists(見積No):
            return 見積No
        
        連番 += 1
        
        # 安全装置（無限ループ防止）
        if 連番 > 999:
            # 999を超えた場合はタイムスタンプを追加
            import time
            timestamp = str(int(time.time()))[-3:]  # 末尾3桁
            return f"{発行日_str}{timestamp}"

def set_customer_selection(顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2):
    """顧客選択を設定（住所細分化対応・修正版）"""
    正規化顧客担当者 = 顧客担当者.replace("　", " ").strip()
    
    st.session_state["選択された顧客会社名"] = 顧客会社名
    st.session_state["選択された顧客部署名"] = 顧客部署名
    st.session_state["選択された顧客担当者"] = 正規化顧客担当者
    st.session_state["選択された郵便番号"] = 郵便番号
    st.session_state["選択された住所1"] = 住所1
    st.session_state["選択された住所2"] = 住所2
    
    # 統合住所も設定（互換性のため）
    統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
    st.session_state["選択された顧客住所"] = 統合住所

def clear_customer_session():
    """顧客セッションをクリア（住所細分化対応・完全版）"""
    keys_to_clear = [
        # 選択された顧客情報
        "選択された顧客会社名", "選択された顧客部署名", "選択された顧客担当者",
        "選択された郵便番号", "選択された住所1", "選択された住所2", "選択された顧客住所",
        # 入力中の顧客情報
        "入力中_顧客会社名選択", "入力中_顧客会社名", "入力中_顧客担当者選択",
        "入力中_顧客担当者", "入力中_顧客部署名", 
        "入力中_郵便番号", "入力中_住所1", "入力中_住所2", "入力中_顧客住所",
        # 補完・編集関連
        "住所補完済み", "編集中顧客"
    ]
    
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]

def clear_customer_input_session():
    """顧客入力中データのみをクリア（選択された顧客情報は保持・完全版）"""
    input_keys_to_clear = [
        # 入力中の顧客情報のみ
        "入力中_顧客会社名選択", "入力中_顧客会社名", "入力中_顧客担当者選択",
        "入力中_顧客担当者", "入力中_顧客部署名", 
        "入力中_郵便番号", "入力中_住所1", "入力中_住所2", "入力中_顧客住所",
        # 補完フラグ
        "住所補完済み"
    ]
    
    for key in input_keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]

# タブ1: 顧客情報入力
def render_customer_tab(顧客一覧_df):
    """顧客情報入力タブを表示（JSONデータ対応修正版）"""
    st.header("② 顧客情報を入力")
    
    # JSONから最新の顧客データを読み込み
    最新顧客データ = load_customers_json()
    
    # DataFrameに変換（既存コードとの互換性のため）
    if 最新顧客データ:
        顧客一覧 = pd.DataFrame(最新顧客データ)
    else:
        顧客一覧 = pd.DataFrame(columns=["顧客No", "顧客会社名", "顧客部署名", "顧客担当者", "顧客住所", "郵便番号", "住所1", "住所2"])
    
    # 編集モードのチェック
    編集中顧客 = st.session_state.get("編集中顧客")
    編集モード = 編集中顧客 is not None
    
    if 編集モード:
        st.info(f"顧客情報を編集中: {編集中顧客['顧客会社名']} - {編集中顧客['顧客担当者']}")
    
    # 登録済みの顧客情報がある場合は表示（編集モードでない場合のみ）
    選択顧客会社名 = st.session_state.get("選択された顧客会社名", "")
    選択顧客部署名 = st.session_state.get("選択された顧客部署名", "")
    選択顧客担当者 = st.session_state.get("選択された顧客担当者", "")
    
    if 選択顧客会社名 and 選択顧客担当者 and not 編集モード:
        st.success("✅ 顧客情報が登録済みです")
        st.write(f"**会社名:** {選択顧客会社名}")
        st.write(f"**部署名:** {選択顧客部署名}")
        st.write(f"**担当者:** {選択顧客担当者}")
        
        # 住所の表示（細分化対応）
        郵便番号 = st.session_state.get("選択された郵便番号", "")
        住所1 = st.session_state.get("選択された住所1", "")
        住所2 = st.session_state.get("選択された住所2", "")
        統合住所 = st.session_state.get("選択された顧客住所", "")
        
        if 郵便番号 or 住所1 or 住所2:
            st.write(f"**住所:** {郵便番号} {住所1} {住所2}")
        elif 統合住所:
            st.write(f"**住所:** {統合住所}")
        else:
            st.write("**住所:** 未設定")
        
        col1, col2 = st.columns(2)
        if col1.button("別の顧客を登録", key="customer_change"):
            # 顧客情報をクリアして新規入力モードに
            clear_customer_session()
            st.rerun()
        
        if col2.button("**案件情報の入力に進む**", type="primary", key="customer_to_project"):
            st.session_state["アクティブタブ"] = "③ 案件情報を入力"
            st.rerun()
        
        return
    
    # 新規入力または編集モードの場合の入力フォーム
    st.subheader("顧客情報の入力")
    
    # 編集モードの場合の初期値設定（旧住所の移行対応）
    if 編集モード:
        初期_顧客会社名 = 編集中顧客["顧客会社名"]
        初期_顧客部署名 = 編集中顧客.get("顧客部署名", "")
        初期_顧客担当者 = 編集中顧客["顧客担当者"]
        
        # 新形式の住所を優先、なければ旧住所を分割
        初期_郵便番号 = 編集中顧客.get("郵便番号", "")
        初期_住所1 = 編集中顧客.get("住所1", "")
        初期_住所2 = 編集中顧客.get("住所2", "")
        
        # 新形式の住所がなく、旧住所がある場合は分割
        if not (初期_郵便番号 or 初期_住所1 or 初期_住所2):
            旧住所 = 編集中顧客.get("顧客住所", "")
            if 旧住所:
                from estimate_excel_writer import parse_address
                初期_郵便番号, 初期_住所1, 初期_住所2 = parse_address(旧住所)
                st.info("旧形式の住所を新形式に変換しました")
    else:
        初期_顧客会社名 = st.session_state.get("入力中_顧客会社名", "")
        初期_顧客部署名 = st.session_state.get("入力中_顧客部署名", "")
        初期_顧客担当者 = st.session_state.get("入力中_顧客担当者", "")
        初期_郵便番号 = st.session_state.get("入力中_郵便番号", "")
        初期_住所1 = st.session_state.get("入力中_住所1", "")
        初期_住所2 = st.session_state.get("入力中_住所2", "")
    
    # 顧客会社名の選択（顧客一覧と同じ順番で表示）
    if not 顧客一覧.empty and "顧客会社名" in 顧客一覧.columns:
        # 顧客一覧と同じ順番で会社名を取得（顧客No順）
        会社名リスト = []
        
        # 顧客一覧をソート（顧客No順、会社名順）
        ソート済み顧客一覧 = 顧客一覧.sort_values(
            by=['顧客No', '顧客会社名'], 
            na_position='last'
        ).reset_index(drop=True)
        
        # 重複を除去しつつ順番を保持
        seen = set()
        for _, row in ソート済み顧客一覧.iterrows():
            会社名 = row.get('顧客会社名', '')
            if 会社名 and 会社名 not in seen:
                会社名リスト.append(会社名)
                seen.add(会社名)
        
        顧客会社名選択肢 = ["（新規入力）"] + 会社名リスト
    else:
        顧客会社名選択肢 = ["（新規入力）"]
    
    # 編集モードまたは既存の選択を復元
    if 編集モード and 初期_顧客会社名 in 顧客会社名選択肢:
        初期index = 顧客会社名選択肢.index(初期_顧客会社名)
    elif not 編集モード:
        前回選択_会社名 = st.session_state.get("入力中_顧客会社名選択", "（新規入力）")
        if 前回選択_会社名 in 顧客会社名選択肢:
            初期index = 顧客会社名選択肢.index(前回選択_会社名)
        else:
            初期index = 0
    else:
        初期index = 0
    
    顧客会社名_選択 = st.selectbox(
        "顧客会社名を選択", 
        顧客会社名選択肢,
        index=初期index,
        key="顧客会社名選択"
    )
    st.session_state["入力中_顧客会社名選択"] = 顧客会社名_選択
    
    if 顧客会社名_選択 == "（新規入力）":
        顧客会社名 = st.text_input(
            "新しい顧客会社名を入力", 
            value=初期_顧客会社名,
            key="新規顧客会社名"
        )
        st.session_state["入力中_顧客会社名"] = 顧客会社名
    else:
        顧客会社名 = 顧客会社名_選択
        st.session_state["入力中_顧客会社名"] = 顧客会社名
    
    # 担当者選択肢の取得（JSONデータから）
    if not 顧客一覧.empty and 顧客会社名 and "顧客担当者" in 顧客一覧.columns:
        会社の顧客 = 顧客一覧[顧客一覧["顧客会社名"] == 顧客会社名]
        if not 会社の顧客.empty:
            担当者候補 = 会社の顧客["顧客担当者"].dropna().unique().tolist()
        else:
            担当者候補 = []
    else:
        担当者候補 = []
    
    担当者選択肢 = ["（新規入力）"] + 担当者候補
    
    # 編集モードまたは既存の選択を復元
    if 編集モード and 初期_顧客担当者 in 担当者選択肢:
        担当者初期index = 担当者選択肢.index(初期_顧客担当者)
    elif not 編集モード:
        前回選択_担当者 = st.session_state.get("入力中_顧客担当者選択", "（新規入力）")
        if 前回選択_担当者 in 担当者選択肢:
            担当者初期index = 担当者選択肢.index(前回選択_担当者)
        else:
            担当者初期index = 0
    else:
        担当者初期index = 0
    
    顧客担当者_選択 = st.selectbox(
        "顧客担当者を選択または入力", 
        担当者選択肢,
        index=担当者初期index,
        key="顧客担当者選択"
    )
    st.session_state["入力中_顧客担当者選択"] = 顧客担当者_選択
    
    if 顧客担当者_選択 == "（新規入力）":
        顧客担当者 = st.text_input(
            "新しい顧客担当者を入力",
            value=初期_顧客担当者,
            key="新規顧客担当者"
        )
        st.session_state["入力中_顧客担当者"] = 顧客担当者
    else:
        顧客担当者 = 顧客担当者_選択
        st.session_state["入力中_顧客担当者"] = 顧客担当者
    
    # 部署名の自動補完（編集モードでない場合）
    if not 編集モード:
        顧客部署補完 = ""
        if 顧客会社名 and 顧客担当者 and not 顧客一覧.empty:
            候補行 = 顧客一覧[
                (顧客一覧["顧客会社名"] == 顧客会社名) & 
                (顧客一覧["顧客担当者"] == 顧客担当者)
            ]
            if not 候補行.empty:
                顧客部署補完 = 候補行.iloc[0].get("顧客部署名", "")
                if 顧客部署補完 and not st.session_state.get("入力中_顧客部署名"):
                    st.info("部署名を自動補完しました")
                    st.session_state["入力中_顧客部署名"] = 顧客部署補完
        
        部署名初期値 = st.session_state.get("入力中_顧客部署名", 顧客部署補完)
    else:
        部署名初期値 = 初期_顧客部署名
    
    顧客部署名 = st.text_input(
        "顧客部署名", 
        value=部署名初期値,
        key="顧客部署名入力"
    )
    st.session_state["入力中_顧客部署名"] = 顧客部署名
    
    # 住所の細分化入力
    st.subheader("住所情報")
    
    # 住所の自動補完（編集モードでない場合のみ）
    if not 編集モード:
        住所補完 = {"郵便番号": "", "住所1": "", "住所2": ""}
        if 顧客会社名 and not 顧客一覧.empty:
            # 顧客一覧から住所情報を取得（旧住所移行対応）
            候補行 = 顧客一覧[顧客一覧["顧客会社名"] == 顧客会社名]
            if not 候補行.empty:
                # 新しい形式の住所情報を優先
                住所補完["郵便番号"] = 候補行.iloc[0].get("郵便番号", "")
                住所補完["住所1"] = 候補行.iloc[0].get("住所1", "")
                住所補完["住所2"] = 候補行.iloc[0].get("住所2", "")
                
                # 旧形式の住所がある場合は分割を試行
                if not any(住所補完.values()):
                    旧住所 = 候補行.iloc[0].get("顧客住所", "")
                    if 旧住所:
                        from estimate_excel_writer import parse_address
                        郵便番号_parsed, 住所1_parsed, 住所2_parsed = parse_address(旧住所)
                        住所補完["郵便番号"] = 郵便番号_parsed
                        住所補完["住所1"] = 住所1_parsed
                        住所補完["住所2"] = 住所2_parsed
                
                # 補完がある場合は情報を表示
                if any(住所補完.values()):
                    # 初回補完時のみメッセージ表示
                    if not st.session_state.get("住所補完済み"):
                        st.info("住所情報を自動補完しました")
                        st.session_state["住所補完済み"] = True
                    
                    # セッション状態への自動設定（既に値がある場合は上書きしない）
                    for key, value in 住所補完.items():
                        if value and not st.session_state.get(f"入力中_{key}"):
                            st.session_state[f"入力中_{key}"] = value
    
    col1, col2 = st.columns(2)
    
    with col1:
        郵便番号 = st.text_input(
            "郵便番号", 
            value=st.session_state.get("入力中_郵便番号", 初期_郵便番号),
            placeholder="例: 123-4567",
            key="郵便番号入力"
        )
        st.session_state["入力中_郵便番号"] = 郵便番号
    
    with col2:
        住所1 = st.text_input(
            "住所1（都道府県・市区町村・番地）", 
            value=st.session_state.get("入力中_住所1", 初期_住所1),
            placeholder="例: 東京都渋谷区神山町9番2号",
            key="住所1入力"
        )
        st.session_state["入力中_住所1"] = 住所1
    
    住所2 = st.text_input(
        "住所2（建物名・階数・号室など）※任意", 
        value=st.session_state.get("入力中_住所2", 初期_住所2),
        placeholder="例: 〇〇ビル3階（建物名がない場合は空欄でOK）",
        key="住所2入力"
    )
    st.session_state["入力中_住所2"] = 住所2
    
    # 同一会社の住所一括更新チェックボックス（編集モードかつ既存会社の場合のみ表示）
    同一会社住所更新 = False
    if 編集モード and not 顧客一覧.empty and 顧客会社名 in 顧客一覧["顧客会社名"].values:
        同一会社住所更新 = st.checkbox(
            f"🏢 同一会社「{顧客会社名}」の全担当者の住所を一括更新する",
            value=False,
            key="同一会社住所更新",
            help="チェックすると、この会社の全担当者の住所が同じ住所に更新されます"
        )
        
        if 同一会社住所更新:
            st.warning("⚠️ この会社の全担当者の住所が更新されます")
    
    st.divider()
    
    # 登録ボタンとナビゲーションボタン
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if 編集モード:
            if st.button("顧客情報を更新する", key="customer_update"):
                if not 顧客会社名 or not 顧客担当者:
                    st.warning("顧客会社名と担当者を入力してください")
                else:
                    success, message = update_customer_in_json(
                        編集中顧客, 顧客会社名, 顧客部署名, 顧客担当者, 
                        郵便番号, 住所1, 住所2, 同一会社住所更新
                    )
                    if success:
                        st.success(message)
                        st.session_state["編集中顧客"] = None
                        clear_customer_session()
                        st.session_state["アクティブタブ"] = "⑤ 顧客一覧"
                        st.rerun()
                    else:
                        st.error(message)
        else:
            if st.button("顧客情報を登録する", key="customer_register"):
                if not 顧客会社名 or not 顧客担当者:
                    st.warning("顧客会社名と担当者を入力してください")
                else:
                    register_customer(顧客一覧, 顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2)
    
    with col2:
        if 編集モード:
            if st.button("編集をキャンセル", key="customer_cancel_edit"):
                st.session_state["編集中顧客"] = None
                clear_customer_session()
                st.session_state["アクティブタブ"] = "⑤ 顧客一覧"
                st.rerun()
        else:
            # 顧客情報が入力されていれば登録なしでも次に進める
            if 顧客会社名 and 顧客担当者:
                if st.button("**案件情報の入力に進む**", type="primary", key="customer_to_project_direct"):
                    # 一時的に顧客情報を設定
                    統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
                    
                    st.session_state["選択された顧客会社名"] = 顧客会社名
                    st.session_state["選択された顧客部署名"] = 顧客部署名
                    st.session_state["選択された顧客担当者"] = 顧客担当者
                    st.session_state["選択された郵便番号"] = 郵便番号
                    st.session_state["選択された住所1"] = 住所1
                    st.session_state["選択された住所2"] = 住所2
                    st.session_state["選択された顧客住所"] = 統合住所
                    
                    st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                    st.rerun()
            else:
                st.button("**案件情報の入力に進む**", disabled=True, help="顧客会社名と担当者を入力してください", key="customer_to_project_disabled")
    
    with col3:
        if 編集モード:
            if st.button("顧客一覧に戻る", key="back_to_customer_list"):
                st.session_state["編集中顧客"] = None
                clear_customer_session()
                st.session_state["アクティブタブ"] = "⑤ 顧客一覧"
                st.rerun()

def register_customer(顧客一覧_df, 顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2):
    """顧客情報をJSONに登録（セッション維持版）"""
    success, message = add_customer_to_json(顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2)
    
    if success:
        # 新規登録成功時
        st.success(f"✅ 顧客情報を登録しました")
        # 顧客情報をセッション状態に設定（統合住所も含む）
        統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
        
        st.session_state["選択された顧客会社名"] = 顧客会社名
        st.session_state["選択された顧客部署名"] = 顧客部署名
        st.session_state["選択された顧客担当者"] = 顧客担当者
        st.session_state["選択された郵便番号"] = 郵便番号
        st.session_state["選択された住所1"] = 住所1
        st.session_state["選択された住所2"] = 住所2
        st.session_state["選択された顧客住所"] = 統合住所
        
        # 入力中データのみをクリア（選択された顧客情報は保持）
        clear_customer_input_session()
        
        # タブは移動しない（ユーザーが手動で移動）
        st.info("💡 下の「案件情報の入力に進む」ボタンで次のステップに進んでください。")
        # st.rerun() を削除 - 自動rerunを回避
    else:
        if "既に登録されています" in message:
            # 既存顧客の場合
            st.info(f"ℹ️ 顧客情報が登録済みです")
            # 既存顧客の場合も情報を設定（統合住所も含む）
            統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
            
            st.session_state["選択された顧客会社名"] = 顧客会社名
            st.session_state["選択された顧客部署名"] = 顧客部署名
            st.session_state["選択された顧客担当者"] = 顧客担当者
            st.session_state["選択された郵便番号"] = 郵便番号
            st.session_state["選択された住所1"] = 住所1
            st.session_state["選択された住所2"] = 住所2
            st.session_state["選択された顧客住所"] = 統合住所
            
            # 入力中データのみをクリア（選択された顧客情報は保持）
            clear_customer_input_session()
            
            # タブは移動しない（ユーザーが手動で移動）
            st.info("💡 下の「案件情報の入力に進む」ボタンで次のステップに進んでください。")
            # st.rerun() を削除 - 自動rerunを回避
        else:
            st.error(f"❌ {message}")
    
    return 顧客一覧_df

def update_customer_in_json(元顧客データ, 顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2, 同一会社住所更新=False):
    """顧客情報をJSONで更新（同一会社住所一括更新対応）"""
    try:
        customers = load_customers_json()
        
        # 元の顧客データを検索
        updated_count = 0
        for i, customer in enumerate(customers):
            if (customer.get("顧客会社名") == 元顧客データ["顧客会社名"] and 
                customer.get("顧客部署名") == 元顧客データ.get("顧客部署名", "") and 
                customer.get("顧客担当者") == 元顧客データ["顧客担当者"]):
                
                # 重複チェック（自分以外）
                if 顧客会社名 != 元顧客データ["顧客会社名"] or 顧客担当者 != 元顧客データ["顧客担当者"]:
                    for j, other_customer in enumerate(customers):
                        if (i != j and 
                            other_customer.get("顧客会社名") == 顧客会社名 and 
                            other_customer.get("顧客部署名") == 顧客部署名 and 
                            other_customer.get("顧客担当者") == 顧客担当者):
                            return False, "同じ顧客情報が既に存在します"
                
                # 会社名が変更された場合の顧客No再計算
                if 顧客会社名 != 元顧客データ["顧客会社名"]:
                    同一会社の顧客 = [c for c in customers if c.get("顧客会社名") == 顧客会社名 and customers.index(c) != i]
                    if 同一会社の顧客:
                        顧客No = 同一会社の顧客[0].get("顧客No", 1)
                    else:
                        existing_companies = list(set([c.get("顧客会社名") for c in customers if customers.index(c) != i]))
                        顧客No = len(existing_companies) + 1
                else:
                    顧客No = customer.get("顧客No", 1)
                
                # 顧客情報を更新（住所細分化対応）
                customers[i]["顧客No"] = 顧客No
                customers[i]["顧客会社名"] = 顧客会社名
                customers[i]["顧客部署名"] = 顧客部署名
                customers[i]["顧客担当者"] = 顧客担当者
                customers[i]["郵便番号"] = 郵便番号
                customers[i]["住所1"] = 住所1
                customers[i]["住所2"] = 住所2
                customers[i]["更新日"] = datetime.date.today().strftime("%Y-%m-%d")
                
                # 旧住所フィールドも新住所で更新（互換性のため）
                統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
                customers[i]["顧客住所"] = 統合住所
                
                updated_count += 1
                break
        
        # 同一会社住所一括更新の処理
        if 同一会社住所更新 and updated_count > 0:
            同一会社顧客数 = 0
            for customer in customers:
                if customer.get("顧客会社名") == 顧客会社名:
                    customer["郵便番号"] = 郵便番号
                    customer["住所1"] = 住所1
                    customer["住所2"] = 住所2
                    customer["更新日"] = datetime.date.today().strftime("%Y-%m-%d")
                    # 旧住所フィールドも更新
                    統合住所 = f"{郵便番号} {住所1} {住所2}".strip()
                    customer["顧客住所"] = 統合住所
                    同一会社顧客数 += 1
        
        if updated_count == 0:
            return False, "更新対象の顧客が見つかりませんでした"
        
        # 顧客Noと会社名で自動並び替え
        customers.sort(key=lambda x: (x.get("顧客No", 999), x.get("顧客会社名", ""), x.get("顧客担当者", "")))
        
        if save_customers_json(customers):
            if 同一会社住所更新:
                return True, f"顧客「{顧客会社名}」を更新しました。同一会社{同一会社顧客数}名の住所も一括更新されました。"
            else:
                return True, f"顧客「{顧客会社名}」を更新しました"
        else:
            return False, "顧客データの保存に失敗しました"
        
    except Exception as e:
        return False, f"更新処理でエラーが発生しました: {e}"

# タブ2: 案件情報入力（JSON専用版）
def search_json_projects(顧客会社名, 顧客部署名, 顧客担当者):
    """JSONファイルから案件を検索（シンプル版）"""
    
    # JSONファイルから該当する案件を検索
    該当案件リスト = []
    
    if not os.path.exists(DATA_FOLDER):
        st.info("該当する案件がありません。新規入力を選択してください。")
        return
    
    try:
        json_files = [f for f in os.listdir(DATA_FOLDER) if f.endswith('.json')]
    except Exception:
        st.info("データフォルダの読み込みに失敗しました。新規入力を選択してください。")
        return
    
    if not json_files:
        st.info("該当する案件がありません。新規入力を選択してください。")
        return
    
    # JSON検索処理
    for file in json_files:
        try:
            ファイルパス = os.path.join(DATA_FOLDER, file)
            with open(ファイルパス, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            if isinstance(data, dict):
                # 顧客情報の取得と正規化
                file_顧客会社名 = data.get("顧客会社名", "").strip()
                file_顧客担当者 = data.get("顧客担当者", "").strip()
                
                # スペースを除去して比較
                入力_顧客会社名_clean = 顧客会社名.replace(" ", "").replace("　", "")
                入力_顧客担当者_clean = 顧客担当者.replace(" ", "").replace("　", "")
                file_顧客会社名_clean = file_顧客会社名.replace(" ", "").replace("　", "")
                file_顧客担当者_clean = file_顧客担当者.replace(" ", "").replace("　", "")
                
                # 一致判定
                会社一致 = file_顧客会社名_clean == 入力_顧客会社名_clean
                担当者一致 = file_顧客担当者_clean == 入力_顧客担当者_clean
                
                if 会社一致 and 担当者一致:
                    # 明細から合計金額を計算
                    明細合計 = sum(item.get("金額", 0) for item in data.get("明細リスト", []))
                    売上額 = data.get("売上額", 0) if data.get("売上額", 0) > 0 else 明細合計
                    
                    該当案件リスト.append({
                        "見積No": data.get("見積No", ""),
                        "案件名": data.get("案件名", "案件名未設定"),
                        "発行日": data.get("発行日", ""),
                        "状況": data.get("状況", "見積中"),
                        "売上額": 売上額
                    })
                    
        except Exception as e:
            continue
    
    # 結果の表示
    if not 該当案件リスト:
        st.info("該当する案件がありません。新規入力を選択してください。")
        return
    
    # 案件を発行日順でソート
    該当案件リスト.sort(key=lambda x: x["発行日"], reverse=True)
    
    # 案件選択リストの作成（詳細情報付き）
    案件表示リスト = []
    案件辞書 = {}
    for 案件 in 該当案件リスト:
        # 表示形式: 案件名 (見積No) - ¥金額 [状況]
        表示名 = f"{案件['案件名']} ({案件['見積No']}) - ¥{案件['売上額']:,} [{案件['状況']}]"
        案件表示リスト.append(表示名)
        案件辞書[表示名] = 案件
    
    # 案件選択
    案件選択 = st.selectbox("案件を選択", [""] + 案件表示リスト, key="json_project_select")

    if 案件選択:
        選択案件 = 案件辞書[案件選択]
        見積No候補 = 選択案件["見積No"]

        # JSON反映ボタン
        if st.button("🔄 JSONを反映する"):
            成功 = auto_load_json_by_estimate_no(見積No候補, auto_rerun=True)
            if 成功:
                st.success(f"{見積No候補} のJSONデータを反映しました")
            else:
                st.error("JSONの読み込みに失敗しました")

def handle_new_project_input():
    """新規案件入力の処理（None対応版）"""
    案件名 = st.text_input("案件名", value=st.session_state.get("案件名", ""))
    st.session_state["案件名"] = 案件名
    
    # 見積番号の生成（発行日がNoneでない場合のみ）
    if not st.session_state.get("見積No"):
        発行日 = st.session_state.get("発行日")
        if 発行日 is not None:  # Noneチェックを追加
            st.session_state["見積No"] = generate_estimate_no(発行日)
        # 発行日がNoneの場合は見積番号も空のまま

def render_project_tab():
    """案件情報入力タブを表示（JSON専用・案件一覧なし・顧客チェック修正版）"""
    st.header("③ 案件情報を入力")
    
    # 選択された顧客情報を表示
    顧客会社名 = st.session_state.get("選択された顧客会社名", "")
    顧客部署名 = st.session_state.get("選択された顧客部署名", "")
    顧客担当者 = st.session_state.get("選択された顧客担当者", "")
    
    # 顧客情報チェックの条件を緩和（会社名と担当者があれば十分）
    if not 顧客会社名 or not 顧客担当者:
        st.warning("顧客情報が不完全です。顧客会社名と担当者を設定してください。")
        if st.button("顧客情報入力に戻る"):
            st.session_state["アクティブタブ"] = "② 顧客情報を入力"
            st.rerun()
        return
    
    # 顧客情報の表示（読み取り専用）
    st.success("✅ 顧客情報が設定済みです")
    
    # 顧客情報を見やすく表示
    customer_col1, customer_col2, customer_col3 = st.columns(3)
    with customer_col1:
        st.text_input("顧客会社名", 顧客会社名, disabled=True, key="display_company")
    with customer_col2:
        st.text_input("顧客部署名", 顧客部署名 or "（未設定）", disabled=True, key="display_dept")
    with customer_col3:
        st.text_input("顧客担当者", 顧客担当者, disabled=True, key="display_contact")
    
    # 顧客情報変更ボタン
    if st.button("🔄 顧客情報を変更", key="change_customer"):
        st.session_state["アクティブタブ"] = "② 顧客情報を入力"
        st.rerun()
    
    st.divider()
    
    # 案件入力方法の選択
    案件入力方法 = st.radio("案件の入力方法", ["新規入力", "既存案件を選択"])
    
    if 案件入力方法 == "既存案件を選択":
        # JSON専用の案件検索（案件一覧を渡さない）
        search_json_projects(顧客会社名, 顧客部署名, 顧客担当者)
        
        # 既存案件を選択した場合でも案件名を編集可能にする
        st.subheader("案件情報の編集")
        案件名 = st.text_input("案件名", value=st.session_state.get("案件名", ""), key="existing_project_name")
        st.session_state["案件名"] = 案件名
    else:
        handle_new_project_input()
    
    # 共通の入力項目（案件一覧を渡さない）
    render_common_project_inputs()

def get_max_sequence_for_date(発行日_str):
    """指定日付の既存見積番号から最大連番を取得"""
    max_sequence = 0
    
    if not os.path.exists(DATA_FOLDER):
        return max_sequence
    
    try:
        json_files = [f for f in os.listdir(DATA_FOLDER) if f.endswith('.json')]
        
        for file in json_files:
            try:
                # ファイル名から見積番号を抽出（拡張子を除く）
                見積No = file.replace('.json', '')
                
                # 見積番号が期待する形式かチェック（YYYYMMDDXXX）
                if len(見積No) == 11 and 見積No[:8] == 発行日_str and 見積No[8:].isdigit():
                    sequence = int(見積No[8:])  # 末尾3桁を取得
                    max_sequence = max(max_sequence, sequence)
                
                # JSONファイル内の発行日もチェック（ファイル名と不一致の場合に備えて）
                ファイルパス = os.path.join(DATA_FOLDER, file)
                with open(ファイルパス, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                if isinstance(data, dict):
                    file_発行日 = data.get("発行日", "")
                    file_見積No = data.get("見積No", "")
                    
                    # 発行日の正規化
                    if isinstance(file_発行日, str):
                        if len(file_発行日) == 10 and "-" in file_発行日:
                            try:
                                file_date = datetime.datetime.strptime(file_発行日, "%Y-%m-%d").date()
                                file_発行日_str = file_date.strftime('%Y%m%d')
                            except:
                                continue
                        elif len(file_発行日) == 8 and file_発行日.isdigit():
                            file_発行日_str = file_発行日
                        else:
                            continue
                        
                        # 同じ発行日で見積番号が正しい形式の場合
                        if (file_発行日_str == 発行日_str and 
                            len(file_見積No) == 11 and 
                            file_見積No[:8] == 発行日_str and 
                            file_見積No[8:].isdigit()):
                            sequence = int(file_見積No[8:])
                            max_sequence = max(max_sequence, sequence)
                            
            except Exception:
                continue
                
    except Exception:
        pass
    
    return max_sequence

def generate_next_estimate_no(発行日):
    """指定日付の次の見積番号を生成（None対応版）"""
    # 発行日がNoneの場合は空文字列を返す
    if 発行日 is None:
        return ""
    
    発行日_str = 発行日.strftime('%Y%m%d')
    
    # 既存の最大連番を取得
    max_sequence = get_max_sequence_for_date(発行日_str)
    
    # 次の連番を生成
    next_sequence = max_sequence + 1
    
    # 見積番号を生成
    見積No = f"{発行日_str}{str(next_sequence).zfill(3)}"
    
    return 見積No

def update_estimate_number_and_overwrite(旧見積No, 新見積No):
    """見積番号を更新して旧ファイルを削除"""
    try:
        旧ファイルパス = os.path.join(DATA_FOLDER, f"{旧見積No}.json")
        新ファイルパス = os.path.join(DATA_FOLDER, f"{新見積No}.json")
        
        # 旧ファイルが存在する場合のみ処理
        if os.path.exists(旧ファイルパス):
            # 旧ファイルを削除
            os.remove(旧ファイルパス)
            return True
        return False
        
    except Exception as e:
        st.error(f"ファイル更新エラー: {e}")
        return False
        
def render_common_project_inputs():
    """共通の案件入力項目を表示（見積番号生成確認対応版・タブ状態保持強化）"""
    
    # 基本情報セクション
    st.subheader("基本情報")
    
    # 見積番号生成確認ダイアログの処理
    見積番号生成確認状態 = st.session_state.get("見積番号生成確認状態")
    if 見積番号生成確認状態:
        st.info("📋 見積番号の設定確認")
        st.write(f"**発行日:** {見積番号生成確認状態['発行日']}")
        st.write(f"**見積No.を {見積番号生成確認状態['新見積No']} に設定します**")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ OK", key="confirm_generate_estimate_no", type="primary"):
                # 見積番号を確定
                st.session_state["見積No"] = 見積番号生成確認状態['新見積No']
                st.session_state["発行日"] = 見積番号生成確認状態['発行日']
                
                # 確認状態をクリア
                del st.session_state["見積番号生成確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.success(f"✅ 見積No.を {見積番号生成確認状態['新見積No']} に設定しました")
                st.rerun()
        
        with col2:
            if st.button("❌ キャンセル", key="cancel_generate_estimate_no"):
                # 発行日を元に戻す
                st.session_state["発行日"] = 見積番号生成確認状態['元発行日']
                
                # 確認状態をクリア
                del st.session_state["見積番号生成確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.info("発行日を選び直してください")
                st.rerun()
        
        # サイドバーを明示的に表示
        render_sidebar_status()
        
        # 確認中は他の入力を無効化
        st.stop()

    # 上書き確認ダイアログの処理（修正版）
    上書き確認状態 = st.session_state.get("上書き確認状態")
    if 上書き確認状態:
        st.error("⚠️ 見積番号の変更確認")
        st.write(f"**発行日変更:** {上書き確認状態['旧発行日']} → {上書き確認状態['新発行日']}")
        st.write(f"**見積番号変更:** {上書き確認状態['旧見積No']} → {上書き確認状態['新見積No']}")
        st.warning("旧データ（{0}）は削除され、新しい見積番号（{1}）で保存されます。".format(
            上書き確認状態['旧見積No'], 上書き確認状態['新見積No']
        ))
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ はい、上書きします", key="confirm_overwrite", type="primary"):
                # 上書き処理を実行
                st.session_state["見積No"] = 上書き確認状態['新見積No']
                st.session_state["発行日"] = 上書き確認状態['新発行日']
                
                # 上書き処理をスケジュール
                st.session_state["上書き処理"] = {
                    "旧見積No": 上書き確認状態['旧見積No'],
                    "新見積No": 上書き確認状態['新見積No'],
                    "実行予定": True
                }
                
                # 確認状態をクリア
                del st.session_state["上書き確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.success(f"見積番号を {上書き確認状態['新見積No']} に変更しました")
                st.info("次回保存時に旧データが削除されます")
                st.rerun()
        
        with col2:
            if st.button("❌ いいえ、元に戻します", key="cancel_overwrite"):
                # 発行日を元に戻す
                st.session_state["発行日"] = 上書き確認状態['旧発行日']
                
                # 確認状態をクリア
                del st.session_state["上書き確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.info("発行日の変更をキャンセルしました")
                st.rerun()
        
        # サイドバーを明示的に表示
        render_sidebar_status()
        
        # 上書き確認中は他の入力を無効化
        st.stop()

    # 上書き確認ダイアログの処理（サイドバー維持版）
    上書き確認状態 = st.session_state.get("上書き確認状態")
    if 上書き確認状態:
        st.error("⚠️ 見積番号の変更確認")
        st.write(f"**発行日変更:** {上書き確認状態['旧発行日']} → {上書き確認状態['新発行日']}")
        st.write(f"**見積番号変更:** {上書き確認状態['旧見積No']} → {上書き確認状態['新見積No']}")
        st.warning("旧データ（{0}）は削除され、新しい見積番号（{1}）で保存されます。".format(
            上書き確認状態['旧見積No'], 上書き確認状態['新見積No']
        ))
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ はい、上書きします", key="confirm_overwrite", type="primary"):
                # 上書き処理を実行
                st.session_state["見積No"] = 上書き確認状態['新見積No']
                st.session_state["発行日"] = 上書き確認状態['新発行日']
                
                # 上書き処理をスケジュール
                st.session_state["上書き処理"] = {
                    "旧見積No": 上書き確認状態['旧見積No'],
                    "新見積No": 上書き確認状態['新見積No'],
                    "実行予定": True
                }
                
                # 確認状態をクリア
                del st.session_state["上書き確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.success(f"見積番号を {上書き確認状態['新見積No']} に変更しました")
                st.info("次回保存時に旧データが削除されます")
                st.rerun()
        
        with col2:
            if st.button("❌ いいえ、元に戻します", key="cancel_overwrite"):
                # 発行日を元に戻す
                st.session_state["発行日"] = 上書き確認状態['旧発行日']
                
                # 確認状態をクリア
                del st.session_state["上書き確認状態"]
                
                # タブ状態を保持してからrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.info("発行日の変更をキャンセルしました")
                st.rerun()
        
        # サイドバーを明示的に表示させてからstop
        render_sidebar_status()
        
        # 上書き確認中は他の入力を無効化
        st.stop()
    
    col1, col2 = st.columns(2)
    
    # 発行日入力（初期値空白対応版・タブ状態保持強化）
    with col1:
        # 初期状態では発行日を空にする（新規案件の場合）
        見積No = st.session_state.get("見積No", "")
        if 見積No:
            # 既存案件の場合は現在の発行日を初期値として使用
            現在の発行日 = st.session_state.get("発行日", datetime.date.today())
        else:
            # 新規案件の場合は初期値をNoneにして空欄表示
            現在の発行日 = st.session_state.get("発行日", None)
        
        発行日 = st.date_input("発行日", value=現在の発行日, key="project_issue_date")
        
        # 発行日が変更された場合の処理（タブ状態保持強化）
        if 発行日 != 現在の発行日:
            旧見積No = st.session_state.get("見積No", "")
            
            if 旧見積No:
                # 既存案件の場合、上書き確認を表示
                新見積No = generate_next_estimate_no(発行日)
                
                if 新見積No and 旧見積No != 新見積No:  # 新見積Noが空でないことを確認
                    # 上書き確認状態を設定
                    st.session_state["上書き確認状態"] = {
                        "旧見積No": 旧見積No,
                        "新見積No": 新見積No,
                        "旧発行日": 現在の発行日,
                        "新発行日": 発行日
                    }
                    # タブ状態を明示的に保持してからrerun
                    st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                    st.rerun()
                else:
                    # 同じ見積番号の場合はそのまま更新
                    st.session_state["発行日"] = 発行日
            else:
                # 新規作成の場合：見積番号生成確認を表示
                新見積No = generate_next_estimate_no(発行日)
                if 新見積No:  # 新見積Noが空でないことを確認
                    st.session_state["見積番号生成確認状態"] = {
                        "発行日": 発行日,
                        "新見積No": 新見積No,
                        "元発行日": 現在の発行日
                    }
                    # タブ状態を明示的に保持してからrerun
                    st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                    st.rerun()
        
        # 新規案件で発行日が未選択の場合は警告を表示
        if not 見積No and not 発行日:
            st.info("💡 発行日を選択すると見積番号が生成されます")
    
    # 見積番号表示（修正版）
    with col2:
        見積No表示 = st.session_state.get("見積No", "")
        
        if 見積No表示:
            # 見積番号表示
            st.text_input("見積No.", value=見積No表示, disabled=True, key="project_estimate_no")
            
            # 上書き予定の表示
            上書き処理 = st.session_state.get("上書き処理")
            if 上書き処理 and 上書き処理.get("実行予定"):
                st.warning(f"⏳ 保存時に {上書き処理['旧見積No']} を削除予定")
            
            # 重複チェック表示
            if check_estimate_no_exists(見積No表示):
                st.error("⚠️ この見積番号は既に存在します")
                # 自動で次の番号を提案
                if st.button("🔄 次の番号を自動生成", key="auto_generate_next"):
                    new_estimate_no = generate_next_estimate_no(st.session_state["発行日"])
                    st.session_state["見積No"] = new_estimate_no
                    st.success(f"新しい見積番号: {new_estimate_no}")
                    # タブ状態を保持してrerun
                    st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                    st.rerun()
            else:
                st.success("✅ 利用可能な見積番号です")
            
            # 手動再生成ボタン
            if st.button("🔄 見積番号を再生成", key="manual_regenerate"):
                new_estimate_no = generate_next_estimate_no(st.session_state["発行日"])
                st.session_state["見積No"] = new_estimate_no
                st.success(f"見積番号を再生成しました: {new_estimate_no}")
                # タブ状態を保持してrerun
                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                st.rerun()
        else:
            # 見積番号が未生成の場合
            st.text_input("見積No.", value="", disabled=True, key="project_estimate_no_empty", placeholder="発行日を選択してください")
            st.info("💡 発行日を選択すると見積番号生成の確認ダイアログが表示されます")
    
    # 発行者名
    発行者名 = st.selectbox(
        "発行者名", 
        options=ISSUER_LIST, 
        index=ISSUER_LIST.index(st.session_state.get("発行者名", ISSUER_LIST[0])) 
        if st.session_state.get("発行者名") in ISSUER_LIST else 0,
        key="project_issuer"
    )
    st.session_state["発行者名"] = 発行者名
    
    # 担当部署
    担当部署選択肢 = ["", "映像制作部", "翻訳制作部", "完プロ制作部", "生字幕制作部", "字幕展開部"]
    担当部署初期index = 0
    現在の担当部署 = st.session_state.get("担当部署", "")
    if 現在の担当部署 in 担当部署選択肢:
        担当部署初期index = 担当部署選択肢.index(現在の担当部署)
    
    担当部署 = st.selectbox(
        "担当部署", 
        options=担当部署選択肢,
        index=担当部署初期index,
        key="project_department_select",
        help="案件全体の担当部署。明細で売上先部署が未設定の場合、この部署が適用されます"
    )
    
    # 担当部署が変更された場合のみセッション状態を更新
    if 担当部署 != st.session_state.get("担当部署", ""):
        st.session_state["担当部署"] = 担当部署
    
    st.divider()
    
    # 進捗管理セクション
    st.subheader("進捗管理")
    
    col1, col2, col3 = st.columns(3)
    
    # 状況
    with col1:
        状況選択肢 = ["見積中", "受注", "納品済", "請求済", "不採用", "失注"] 
        現在の状況 = st.session_state.get("状況", "見積中")
        # 状況がリストにない場合はデフォルト値を使用
        状況初期index = 0
        if 現在の状況 in 状況選択肢:
            状況初期index = 状況選択肢.index(現在の状況)
        
        状況 = st.selectbox(
            "状況", 
            options=状況選択肢,
            index=状況初期index,
            key="project_status"
        )
        st.session_state["状況"] = 状況
    
    # 受注日（状況に応じて有効/無効）
    with col2:
        受注日無効 = 状況 in ["見積中", "失注"]
        if 受注日無効:
            st.date_input("受注日", value=None, disabled=True, key="project_order_date_disabled")
            st.session_state["受注日"] = None
        else:
            受注日 = st.date_input(
                "受注日", 
                value=st.session_state.get("受注日"), 
                key="project_order_date"
            )
            st.session_state["受注日"] = 受注日
    
    # 納品日（常に入力可能）
    with col3:
        納品日 = st.date_input(
            "納品日", 
            value=st.session_state.get("納品日"), 
            key="project_delivery_date"
        )
        st.session_state["納品日"] = 納品日
    
    st.divider()
    
    # 金額管理セクション
    st.subheader("金額管理")
    
    # 明細からの合計金額を計算
    明細合計 = sum(item.get("金額", 0) for item in st.session_state.get("明細リスト", []) if not item.get("分類", False))
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 売上額自動更新のチェックボックス
        売上額自動更新 = st.checkbox(
            "明細に応じて売上額を自動更新", 
            value=st.session_state.get("売上額自動更新", True), 
            key="project_auto_update"
        )
        st.session_state["売上額自動更新"] = 売上額自動更新
        
        # 売上額（自動更新または手動入力）
        if 売上額自動更新:
            st.number_input(
                "売上額", 
                value=int(明細合計),  # int()で整数に変換
                disabled=True, 
                format="%d",
                key="project_sales_auto"
            )
            st.session_state["売上額"] = 明細合計
        else:
            # 現在の売上額を安全に取得・変換
            現在の売上額 = st.session_state.get("売上額", 明細合計)
            try:
                現在の売上額 = int(float(現在の売上額)) if 現在の売上額 not in ["", None] else int(明細合計)
            except (ValueError, TypeError):
                現在の売上額 = int(明細合計)
            
            売上額 = st.number_input(
                "売上額", 
                value=現在の売上額, 
                min_value=0, 
                step=1000,
                format="%d",
                key="project_sales_manual"
            )
            st.session_state["売上額"] = 売上額
    
    with col2:
        # 仕入額（安全な変換）
        現在の仕入額 = st.session_state.get("仕入額", 0)
        try:
            現在の仕入額 = int(float(現在の仕入額)) if 現在の仕入額 not in ["", None] else 0
        except (ValueError, TypeError):
            現在の仕入額 = 0
        
        仕入額 = st.number_input(
            "仕入額", 
            value=現在の仕入額, 
            min_value=0, 
            step=1000,
            format="%d",
            key="project_cost"
        )
        st.session_state["仕入額"] = 仕入額
        
        # 粗利と粗利率の自動計算・表示
        売上額_値 = st.session_state.get("売上額", 0)
        try:
            売上額_値 = int(float(売上額_値)) if 売上額_値 not in ["", None] else 0
        except (ValueError, TypeError):
            売上額_値 = 0
        
        粗利 = 売上額_値 - 仕入額
        粗利率 = (粗利 / 売上額_値 * 100) if 売上額_値 > 0 else 0
        
        st.session_state["粗利"] = 粗利
        st.session_state["粗利率"] = 粗利率
        
        # 粗利の表示（色分け）
        if 粗利 >= 0:
            st.success(f"粗利: ¥{粗利:,}")
        else:
            st.error(f"粗利: ¥{粗利:,}")
        
        # 粗利率の表示
        st.info(f"粗利率: {粗利率:.1f}%")
    
    st.divider()
    
    # メモ（タブ②用）
    st.subheader("メモ")
    メモ = st.text_area("メモ", value=st.session_state.get("メモ", ""), key="project_memo", height=100)
    st.session_state["メモ"] = メモ
    
    st.divider()
    
    # 案件情報保存ボタンと明細入力に進むボタン
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])

    with col1:
        if st.button("案件情報を保存", key="save_project_info"):
            # バリデーション
            if not st.session_state.get("見積No"):
                st.warning("発行日を選択して見積番号を生成してください")
            elif not st.session_state.get("納品日"):
                st.warning("納品日を入力してください")
            else:
                # JSON保存（メイン）
                保存データ = {
                    "明細リスト": st.session_state.get("明細リスト", []),
                    "備考": st.session_state.get("備考", "")
                }
                
                if save_meisai_as_json(st.session_state["見積No"], 保存データ):
                    st.success("💾 案件情報をJSONに保存しました")
                else:
                    st.error("❌ 案件情報の保存に失敗しました")

    with col3:
        if st.button("**明細入力に進む**", type="primary", key="project_to_detail"):
            # バリデーション
            if not st.session_state.get("案件名"):
                st.warning("案件名を入力してください")
            elif not st.session_state.get("見積No"):
                st.warning("発行日を選択して見積番号を生成してください")
            elif not st.session_state.get("納品日"):
                st.warning("納品日を入力してください")
            else:
                st.session_state["アクティブタブ"] = "④ 明細情報を入力"
                st.rerun()

def is_language_specified(product_name):
    """商品名に言語が既に指定されているかを判定"""
    import re
    
    # 括弧内の内容を抽出
    pattern = r'（(.+?)）'
    matches = re.findall(pattern, product_name)
    
    if not matches:
        return False
    
    # 言語指定として有効な内容かチェック
    language_patterns = [
        # 翻訳言語パターン（○○→○○）
        r'.+→.+',
        # 単一言語パターン
        r'^(英語|中国語|韓国語|タイ語|ベトナム語|フランス語|ドイツ語|スペイン語|ポルトガル語|イタリア語|ロシア語|アラビア語|ヒンディー語|インドネシア語|マレー語|日本語).*$'
    ]
    
    for match in matches:
        for pattern in language_patterns:
            if re.match(pattern, match):
                return True
    
    return False

def is_translation_product(product_name):
    """翻訳関連商品かどうかを判定（既に言語指定済みは除外）"""
    # 既に言語が指定されている場合は対象外
    if is_language_specified(product_name):
        return False
    
    translation_keywords = [
        "字幕翻訳", "文書翻訳", "通訳", "同時通訳", "逐次通訳"
    ]
    return any(keyword in product_name for keyword in translation_keywords)

def is_single_language_product(product_name):
    """単一言語指定商品かどうかを判定（既に言語指定済みは除外）"""
    # 既に言語が指定されている場合は対象外
    if is_language_specified(product_name):
        return False
    
    single_language_keywords = [
        "翻訳準備費", "言語監修", "編集者", "文字起こし", "SRT作成", 
        "ナレーター派遣", "編集者派遣"
    ]
    return any(keyword in product_name for keyword in single_language_keywords)

def get_language_options():
    """言語選択肢を取得（翻訳用：○○→○○形式）"""
    return {
        "": "言語を選択",
        "日本語→英語": "日本語→英語", 
        "日本語→中国語(簡体字)": "日本語→中国語(簡体字)",
        "日本語→中国語(繁体字)": "日本語→中国語(繁体字)",
        "日本語→韓国語": "日本語→韓国語",
        "日本語→タイ語": "日本語→タイ語",
        "日本語→ベトナム語": "日本語→ベトナム語",
        "日本語→フランス語": "日本語→フランス語",
        "日本語→ドイツ語": "日本語→ドイツ語",
        "日本語→スペイン語": "日本語→スペイン語",
        "英語→日本語": "英語→日本語",
        "中国語(簡体字)→日本語": "中国語(簡体字)→日本語",
        "中国語(繁体字)→日本語": "中国語(繁体字)→日本語",
        "韓国語→日本語": "韓国語→日本語",
        "タイ語→日本語": "タイ語→日本語",
        "ベトナム語→日本語": "ベトナム語→日本語",
        "フランス語→日本語": "フランス語→日本語",
        "ドイツ語→日本語": "ドイツ語→日本語",
        "スペイン語→日本語": "スペイン語→日本語",
        "その他": "その他（手入力）"
    }

def get_single_language_options():
    """単一言語選択肢を取得"""
    return {
        "": "言語を選択",
        "英語": "英語",
        "中国語(簡体字)": "中国語(簡体字)",
        "中国語(繁体字)": "中国語(繁体字)",
        "韓国語": "韓国語",
        "タイ語": "タイ語",
        "ベトナム語": "ベトナム語",
        "フランス語": "フランス語",
        "ドイツ語": "ドイツ語",
        "スペイン語": "スペイン語",
        "ポルトガル語": "ポルトガル語",
        "イタリア語": "イタリア語",
        "ロシア語": "ロシア語",
        "アラビア語": "アラビア語",
        "ヒンディー語": "ヒンディー語",
        "インドネシア語": "インドネシア語",
        "マレー語": "マレー語",
        "その他": "その他（手入力）"
    }

def render_language_selection(base_product_name, current_language="", key_suffix=""):
    """言語選択UIを表示（翻訳商品・単一言語商品両対応）"""
    # 翻訳商品の判定
    if is_translation_product(base_product_name):
        st.info(f"💡 「{base_product_name}」は翻訳関連商品です。言語ペアを指定できます。")
        
        language_options = get_language_options()
        language_keys = list(language_options.keys())
        
        # 現在の言語の初期インデックスを取得
        initial_index = 0
        if current_language and current_language in language_keys:
            initial_index = language_keys.index(current_language)
        elif current_language and current_language not in language_keys:
            # カスタム言語の場合は「その他」を選択
            initial_index = language_keys.index("その他") if "その他" in language_keys else 0
        
        selected_language = st.selectbox(
            "言語ペアを選択", 
            options=language_keys,
            format_func=lambda x: language_options[x],
            index=initial_index,
            key=f"language_select_{key_suffix}"
        )
        
        # カスタム言語入力（その他を選択時）
        final_language = selected_language
        if selected_language == "その他":
            # 現在の言語がカスタム言語の場合は初期値として設定
            initial_custom = ""
            if current_language and current_language not in language_keys:
                initial_custom = current_language
            
            custom_language = st.text_input(
                "カスタム言語ペアを入力", 
                value=initial_custom,
                placeholder="例: タイ語→日本語",
                key=f"custom_language_{key_suffix}"
            )
            if custom_language:
                final_language = custom_language
            else:
                final_language = ""  # カスタム入力が空の場合
        
        # 商品名の構築
        if final_language and final_language != "":
            final_product_name = f"{base_product_name}（{final_language}）"
        else:
            final_product_name = base_product_name
        
        return final_product_name, final_language
    
    # 単一言語商品の判定
    elif is_single_language_product(base_product_name):
        st.info(f"💡 「{base_product_name}」は言語指定可能な商品です。対象言語を指定できます。")
        
        language_options = get_single_language_options()
        language_keys = list(language_options.keys())
        
        # 現在の言語の初期インデックスを取得
        initial_index = 0
        if current_language and current_language in language_keys:
            initial_index = language_keys.index(current_language)
        elif current_language and current_language not in language_keys:
            # カスタム言語の場合は「その他」を選択
            initial_index = language_keys.index("その他") if "その他" in language_keys else 0
        
        selected_language = st.selectbox(
            "対象言語を選択", 
            options=language_keys,
            format_func=lambda x: language_options[x],
            index=initial_index,
            key=f"single_language_select_{key_suffix}"
        )
        
        # カスタム言語入力（その他を選択時）
        final_language = selected_language
        if selected_language == "その他":
            # 現在の言語がカスタム言語の場合は初期値として設定
            initial_custom = ""
            if current_language and current_language not in language_keys:
                initial_custom = current_language
            
            custom_language = st.text_input(
                "カスタム言語を入力", 
                value=initial_custom,
                placeholder="例: ポルトガル語",
                key=f"custom_single_language_{key_suffix}"
            )
            if custom_language:
                final_language = custom_language
            else:
                final_language = ""  # カスタム入力が空の場合
        
        # 商品名の構築
        if final_language and final_language != "":
            final_product_name = f"{base_product_name}（{final_language}）"
        else:
            final_product_name = base_product_name
        
        return final_product_name, final_language
    
    # 言語指定不要な商品
    else:
        return base_product_name, current_language

def extract_base_product_and_language(product_name):
    """商品名から基本商品名と言語情報を分離（翻訳・単一言語両対応）"""
    import re
    
    # 言語パターンにマッチする場合
    pattern = r'^(.+?)（(.+?)）$'
    match = re.match(pattern, product_name)
    
    if match:
        base_name = match.group(1)
        language = match.group(2)
        
        # 翻訳言語として有効かチェック（○○→○○形式）
        translation_language_options = get_language_options()
        if language in translation_language_options or "→" in language:
            return base_name, language
        
        # 単一言語として有効かチェック
        single_language_options = get_single_language_options()
        if language in single_language_options:
            return base_name, language
        
        # その他のカスタム言語
        return base_name, language
    
    # マッチしない場合は全体を基本商品名として扱う
    return product_name, ""

# この位置に追加（extract_base_product_and_language関数の直後）

def is_management_fee_product(product_name):
    """管理費関連商品かどうかを判定"""
    management_fee_keywords = [
        "管理費", "手数料", "事務手数料", "システム利用料", "処理手数料"
    ]
    return any(keyword in product_name for keyword in management_fee_keywords)

def get_percentage_options():
    """％選択肢を取得"""
    return {
        "": "％を選択",
        "5": "5%",
        "10": "10%", 
        "15": "15%",
        "20": "20%",
        "25": "25%",
        "30": "30%",
        "その他": "その他（手入力）"
    }

def render_percentage_selection(base_product_name, current_percentage="", key_suffix=""):
    """％選択UIを表示"""
    if is_management_fee_product(base_product_name):
        st.info(f"💡 「{base_product_name}」は管理費関連商品です。％指定で自動計算できます。")
        
        percentage_options = get_percentage_options()
        percentage_keys = list(percentage_options.keys())
        
        # 現在の％の初期インデックスを取得
        initial_index = 0
        if current_percentage and current_percentage in percentage_keys:
            initial_index = percentage_keys.index(current_percentage)
        elif current_percentage and current_percentage not in percentage_keys:
            # カスタム％の場合は「その他」を選択
            initial_index = percentage_keys.index("その他") if "その他" in percentage_keys else 0
        
        selected_percentage = st.selectbox(
            "管理費％を選択", 
            options=percentage_keys,
            format_func=lambda x: percentage_options[x],
            index=initial_index,
            key=f"percentage_select_{key_suffix}"
        )
        
        # カスタム％入力（その他を選択時）
        final_percentage = selected_percentage
        if selected_percentage == "その他":
            # 現在の％がカスタム％の場合は初期値として設定
            initial_custom = ""
            if current_percentage and current_percentage not in percentage_keys:
                initial_custom = current_percentage
            
            custom_percentage = st.text_input(
                "カスタム％を入力", 
                value=initial_custom,
                placeholder="例: 12.5",
                key=f"custom_percentage_{key_suffix}"
            )
            if custom_percentage:
                final_percentage = custom_percentage
            else:
                final_percentage = ""  # カスタム入力が空の場合
        
        # 商品名の構築
        if final_percentage and final_percentage != "":
            final_product_name = f"{base_product_name}（{final_percentage}%）"
        else:
            final_product_name = base_product_name
        
        return final_product_name, final_percentage
    else:
        return base_product_name, current_percentage

def extract_base_product_and_percentage(product_name):
    """商品名から基本商品名と％情報を分離"""
    import re
    
    # ％パターンにマッチする場合
    pattern = r'^(.+?)（(.+?)%）$'
    match = re.match(pattern, product_name)
    
    if match:
        base_name = match.group(1)
        percentage = match.group(2)
        return base_name, percentage
    
    # マッチしない場合は全体を基本商品名として扱う
    return product_name, ""

def calculate_management_fee_amount(明細リスト, current_index, percentage_str):
    """管理費商品の金額を計算（上位商品の合計×％・分類項目除外版）"""
    try:
        # ％を数値に変換
        percentage = float(percentage_str)
        
        # 現在の商品より上にある商品の合計を計算（分類項目は除外）
        上位商品合計 = 0
        for i in range(current_index):
            if i < len(明細リスト):
                item = 明細リスト[i]
                # 分類項目でない場合のみ合計に含める
                if not item.get("分類", False):
                    金額 = item.get("金額", 0)
                    if isinstance(金額, (int, float)):
                        上位商品合計 += 金額
        
        # ％計算
        管理費金額 = int(上位商品合計 * percentage / 100)
        
        return 管理費金額, 上位商品合計
        
    except (ValueError, TypeError):
        return 0, 0

# タブ3: 明細情報入力（品名プルダウン機能追加・自動補完タイミング修正）
def render_detail_tab(品名一覧):
    """明細情報入力タブを表示（係数機能対応版）"""
    st.header("④ 明細情報を入力")

    # 案件情報が不完全な場合は警告を表示し、タブ②へ戻すボタンを表示
    if not st.session_state.get("案件名") or not st.session_state.get("見積No"):
        st.warning("案件情報が不完全です。タブ③で案件情報を入力してください。")
        if st.button("案件情報入力に戻る", key="detail_back_to_project"):
            st.session_state["アクティブタブ"] = "③ 案件情報を入力"
            st.rerun()
        st.stop()

    # 係数機能の設定
    st.subheader("🔧 表示設定")
    col1, col2 = st.columns([1, 3])
    
    with col1:
        係数機能使用 = st.checkbox(
            "係数機能を使用", 
            value=st.session_state.get("係数機能使用", False),
            help="チェックすると明細に係数列が表示されます"
        )
        st.session_state["係数機能使用"] = 係数機能使用
    
    with col2:
        if 係数機能使用:
            st.info("💡 係数機能が有効です。明細に係数列が表示されます。")
        else:
            st.info("💡 係数機能が無効です。通常の明細表示です。")
    
    st.divider()

    # 最新の商品データを取得（JSONから直接読み込み）
    最新品名一覧 = load_products_json()
    
    # DataFrameに変換（互換性のため）
    if 最新品名一覧:
        最新品名一覧_df = pd.DataFrame(最新品名一覧)
    else:
        最新品名一覧_df = pd.DataFrame(columns=["品名", "単位", "単価", "備考"])

    # 明細一覧の初期化・検証
    if "明細リスト" not in st.session_state:
        st.session_state["明細リスト"] = []
    elif not isinstance(st.session_state["明細リスト"], list):
        st.session_state["明細リスト"] = []

    # 明細一覧の直接編集表示（係数対応版）
    render_editable_detail_list_with_coefficient(最新品名一覧_df)

    # 新規明細追加フォーム（係数対応版）
    render_new_detail_form(最新品名一覧_df)

    # 備考欄
    st.divider()
    st.subheader("見積書全体の備考")
    備考 = st.text_area("備考", value=st.session_state.get("備考", ""), key="estimate_remarks", height=100)
    st.session_state["備考"] = 備考

    st.divider()

    # ボタン群（明細があるときのみ）
    if st.session_state.get("明細リスト"):
        col1, col2, col3 = st.columns([1, 1, 1])
        if col1.button("明細データを保存", type="primary", key="detail_save"):
            save_detail_data()
        if col2.button("見積書を出力", key="detail_export"):
            export_estimate()
        if col3.button("**案件一覧に戻る**", key="detail_to_list"):
            st.success("見積作成が完了しました")
            st.session_state["アクティブタブ"] = "① 案件一覧"
            st.rerun()
    else:
        st.info("明細を追加してから保存してください")

def render_new_detail_form(品名一覧_df):
    """新規明細追加フォーム（係数対応版）"""
    st.subheader("➕ 新規明細を追加")
    係数機能使用 = st.session_state.get("係数機能使用", False)
    
    # 分類（カテゴリ）追加ボタン
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**明細項目の追加**")
    with col2:
        if st.button("📁 分類を追加", key="add_category_button", help="「▼準備・管理」のような分類見出しを追加"):
            st.session_state["分類追加モード"] = True
            st.rerun()
    
    # 分類追加モード
    if st.session_state.get("分類追加モード", False):
        st.info("📁 分類見出しを追加")
        with st.form("分類追加フォーム"):
            分類名 = st.text_input(
                "分類名を入力", 
                placeholder="例: ▼準備・管理、▼本番当日、▼オプション",
                key="category_name_input"
            )
            
            col1, col2 = st.columns(2)
            with col1:
                if st.form_submit_button("✅ 分類を追加", type="primary"):
                    if 分類名:
                        # 分類項目を明細リストに追加
                        分類明細 = {
                            "品名": 分類名,
                            "数量": 0,
                            "単位": "",
                            "係数": 1,  # 係数項目を追加
                            "単価": 0,
                            "金額": 0,
                            "備考": "",
                            "売上先部署": "",
                            "分類": True
                        }
                        
                        if not isinstance(st.session_state["明細リスト"], list):
                            st.session_state["明細リスト"] = []
                        st.session_state["明細リスト"].append(分類明細)
                        
                        st.success(f"✅ 分類「{分類名}」を追加しました")
                        st.session_state["分類追加モード"] = False
                        st.rerun()
                    else:
                        st.warning("分類名を入力してください")
            
            with col2:
                if st.form_submit_button("❌ キャンセル"):
                    st.session_state["分類追加モード"] = False
                    st.rerun()
        
        st.divider()
    
    # 商品選択と反映
    最新商品データ = load_products_json()
    品名候補 = ["（新規入力）"]
    
    if 最新商品データ:
        try:
            品名リスト = [商品.get("品名", "") for 商品 in 最新商品データ if 商品.get("品名")]
            品名候補.extend(品名リスト)
        except:
            pass

    col1, col2 = st.columns([3, 1])
    
    with col1:
        品名選択 = st.selectbox("基本商品名を選択", options=品名候補, key="新規品名選択")
    
    with col2:
        # 反映ボタン
        反映可能 = 品名選択 != "（新規入力）" and 最新商品データ
        if st.button("🔄 反映", disabled=not 反映可能, help="品名一覧の情報を下記フィールドに反映します", key="新規反映ボタン"):
            try:
                該当商品 = None
                for 商品 in 最新商品データ:
                    if 商品.get("品名") == 品名選択:
                        該当商品 = 商品
                        break
                
                if 該当商品:
                    補完情報 = {}
                    
                    if 該当商品.get("単位"):
                        st.session_state["新規反映_単位"] = 該当商品["単位"]
                        補完情報["単位"] = 該当商品["単位"]
                    if 該当商品.get("単価") and 該当商品["単価"] > 0:
                        st.session_state["新規反映_単価"] = int(該当商品["単価"])
                        補完情報["単価"] = f"¥{int(該当商品['単価']):,}"
                    if 該当商品.get("備考"):
                        st.session_state["新規反映_備考"] = 該当商品["備考"]
                        補完情報["備考"] = 該当商品["備考"]
                    
                    if 補完情報:
                        補完内容 = " | ".join([f"{k}: {v}" for k, v in 補完情報.items()])
                        st.success(f"🔄 品名一覧から情報を反映しました: {補完内容}")
                        st.rerun()
                    else:
                        st.warning("この品名には追加情報が登録されていません")
                else:
                    st.error("品名一覧で該当する品名が見つかりませんでした")
            except Exception as e:
                st.error(f"情報の反映中にエラーが発生しました: {e}")

    # 基本品名の決定
    if 品名選択 == "（新規入力）":
        基本品名 = st.text_input("新しい品名を入力", key="新規品名入力")
    else:
        基本品名 = 品名選択

    # 言語選択と％選択の統合処理
    最終品名 = 基本品名
    選択言語 = ""
    選択パーセンテージ = ""
    
    # 翻訳関連商品の判定
    if is_translation_product(基本品名):
        最終品名, 選択言語 = render_language_selection(基本品名, "", "新規")
    # 単一言語商品の判定
    elif is_single_language_product(基本品名):
        最終品名, 選択言語 = render_language_selection(基本品名, "", "新規")
    # 管理費商品の判定
    elif is_management_fee_product(基本品名):
        最終品名, 選択パーセンテージ = render_percentage_selection(基本品名, "", "新規")

    # 新規明細入力フォーム（係数対応版）
    with st.form("新規明細入力フォーム", clear_on_submit=False):
        # 最終的な品名を表示（読み取り専用）
        if 最終品名 != 基本品名:
            st.text_input("最終的な品名", value=最終品名, disabled=True, key="新規最終品名表示")

        if 係数機能使用:
            col1, col2, col3 = st.columns(3)
        else:
            col1, col2 = st.columns(2)
        
        with col1:
            数量 = st.number_input("数量", min_value=0, value=1, step=1, key="新規数量")
            
            単位選択肢 = ["式", "分", "文字", "半日", "日", "時間", "名", "本", "件", "回", "枚", "個"]
            初期単位 = st.session_state.get("新規反映_単位", "式")
            単位初期index = 0
            if 初期単位 in 単位選択肢:
                単位初期index = 単位選択肢.index(初期単位)
            
            単位 = st.selectbox("単位", options=単位選択肢, index=単位初期index, key="新規単位")
        
        if 係数機能使用:
            with col2:
                係数 = st.number_input("係数", min_value=0.5, value=1.0, step=0.5, key="新規係数")
        
        with col2 if not 係数機能使用 else col3:
            # 管理費商品の場合は単価入力を無効化し、％表示
            if 選択パーセンテージ:
                st.text_input("単価", value=f"{選択パーセンテージ}%", disabled=True, key="新規管理費単価表示")
                実際の単価 = 0  # 管理費の場合は単価は0（後で％計算）
            else:
                初期単価 = st.session_state.get("新規反映_単価", 0)
                実際の単価 = st.number_input("単価", value=初期単価, step=100, help="割引の場合はマイナス値を入力してください", key="新規単価")
            
            売上先部署選択肢 = ["", "映像制作部", "翻訳制作部", "完プロ制作部", "生字幕制作部", "字幕展開部"]
            売上先部署 = st.selectbox("売上先部署", options=売上先部署選択肢, key="新規売上先部署")

        初期備考 = st.session_state.get("新規反映_備考", "")
        備考 = st.text_input("備考", value=初期備考, key="新規備考")

        # 金額計算（係数対応版）
        if 選択パーセンテージ:
            # 管理費の場合：現在の明細リストから上位商品の合計を計算
            現在の明細リスト = st.session_state.get("明細リスト", [])
            次のindex = len(現在の明細リスト)  # 新規追加時は最後のindex
            管理費単価, 上位合計 = calculate_management_fee_amount(現在の明細リスト, 次のindex, 選択パーセンテージ)
            実際の単価 = 管理費単価  # 計算結果を単価として使用
            if 係数機能使用:
                金額 = 数量 * 係数 * 実際の単価
            else:
                金額 = 数量 * 実際の単価
            
            st.write(f"**📊 上位商品合計: ¥{上位合計:,}**")
            st.write(f"**💰 計算された単価: ¥{実際の単価:,} ({選択パーセンテージ}%)**")
            if 係数機能使用:
                st.write(f"**💰 合計金額: ¥{金額:,} (数量{数量} × 係数{係数} × 単価¥{実際の単価:,})**")
            else:
                st.write(f"**💰 合計金額: ¥{金額:,} (数量{数量} × 単価¥{実際の単価:,})**")
        else:
            # 通常商品の場合
            if 係数機能使用:
                金額 = 数量 * 係数 * 実際の単価
                st.write(f"**💰 合計金額: ¥{金額:,} (数量{数量} × 係数{係数} × 単価¥{実際の単価:,})**")
            else:
                金額 = 数量 * 実際の単価
                st.write(f"**💰 合計金額: ¥{金額:,} (数量{数量} × 単価¥{実際の単価:,})**")

        # 追加ボタン
        submitted = st.form_submit_button("✅ 明細を追加", type="primary")
        
        if submitted:
            if not 最終品名:
                st.warning("品名を入力してください")
            else:
                新明細 = {
                    "品名": 最終品名,
                    "数量": 数量,
                    "単位": 単位,
                    "係数": 係数 if 係数機能使用 else 1,  # 係数機能使用時のみ係数を設定
                    "単価": 実際の単価,
                    "金額": 金額,
                    "備考": 備考,
                    "売上先部署": 売上先部署,
                    "分類": False
                }
                
                # 管理費商品の場合は特別な情報を追加（デバッグ用）
                if 選択パーセンテージ:
                    新明細["管理費パーセンテージ"] = 選択パーセンテージ
                    新明細["管理費ベース金額"] = 上位合計
                
                try:
                    # 新規品名の場合の商品登録処理（基本商品名のみ登録）
                    if 品名選択 == "（新規入力）" and 基本品名:
                        登録商品名 = 基本品名  # 言語・％情報は含めない
                        success, message = add_product_to_json(登録商品名, 単位, 実際の単価, 備考)
                        if success:
                            if is_management_fee_product(基本品名):
                                st.info(f"✅ 基本商品「{登録商品名}」を商品一覧に登録しました（％情報は明細のみ）")
                            elif is_translation_product(基本品名) or is_single_language_product(基本品名):
                                st.info(f"✅ 基本商品「{登録商品名}」を商品一覧に登録しました（言語情報は明細のみ）")
                            else:
                                st.info(message)
                    
                    # 明細リストに追加
                    if not isinstance(st.session_state["明細リスト"], list):
                        st.session_state["明細リスト"] = []
                    st.session_state["明細リスト"].append(新明細)
                    st.success("✅ 明細を追加しました")

                    # 反映データをクリア
                    for key in ["新規反映_単位", "新規反映_単価", "新規反映_備考"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    
                    st.rerun()

                except Exception as e:
                    st.error(f"明細の追加中にエラーが発生しました: {e}")

def render_detail_summary():
    """明細合計と部署別集計の表示（分類項目除外版）"""
    明細リスト = st.session_state["明細リスト"]
    
    if not 明細リスト:
        return
    
    # 分類項目を除外してから集計
    通常明細リスト = [row for row in 明細リスト if not row.get("分類", False)]
    
    if not 通常明細リスト:
        return
    
    # 合計金額の計算
    合計金額 = sum(row["金額"] for row in 通常明細リスト if isinstance(row["金額"], (int, float)))
    
    # 案件の担当部署を取得
    案件担当部署 = st.session_state.get("担当部署", "")
    
    # 部署別集計の計算（売上先部署が未設定の場合は担当部署を使用）
    部署別集計 = {}
    for row in 通常明細リスト:
        売上先部署 = row.get("売上先部署", "")
        # 売上先部署が未設定の場合は案件の担当部署を使用
        使用部署 = 売上先部署 if 売上先部署 else 案件担当部署
        
        if 使用部署:  # 部署が設定されている場合のみ
            金額 = row["金額"] if isinstance(row["金額"], (int, float)) else 0
            if 使用部署 in 部署別集計:
                部署別集計[使用部署] += 金額
            else:
                部署別集計[使用部署] = 金額
    
    # 合計金額の表示
    合計表示 = f"**合計金額: ¥{合計金額:,}**"

    # 部署別集計を追加
    if 部署別集計:
        部署別表示リスト = []
        for 部署, 金額 in sorted(部署別集計.items()):
            部署別表示リスト.append(f"{部署}: ¥{金額:,}")
        
        if 部署別表示リスト:
            合計表示 += f"　　（{' | '.join(部署別表示リスト)}）"
    
    st.markdown(合計表示)

def add_to_hinmei_list(品名, 単位, 単価, 備考):
    """新しい品名を品名一覧に追加"""
    try:
        # 現在のデータを読み込み
        顧客一覧, 案件一覧, 品名一覧 = load_excel_data(EXCEL_FILENAME)
        
        # 重複チェック
        if not 品名一覧.empty and 品名 in 品名一覧["品名"].values:
            st.info(f"品名「{品名}」は既に商品一覧に登録済みです")
            return
        
        # 新しい品名を追加
        新規品名 = {
            "品名": 品名,
            "単位": 単位,
            "単価": 単価,
            "備考": 備考
        }
        
        品名一覧 = pd.concat([品名一覧, pd.DataFrame([新規品名])], ignore_index=True)
        品名一覧 = 品名一覧.drop_duplicates(subset=["品名"], keep="last").reset_index(drop=True)
        
        if save_to_excel(品名一覧, "品名一覧"):
            st.success(f"🛍️ 新商品「{品名}」を商品一覧に追加しました")
        else:
            st.error("商品一覧への追加に失敗しました")
            
    except Exception as e:
        st.error(f"商品一覧追加エラー: {e}")

def save_detail_data():
    """明細データを保存"""
    if not st.session_state.get("見積No"):
        st.error("見積番号が設定されていません")
        return False
    
    保存データ = {
        "明細リスト": st.session_state["明細リスト"],
        "備考": st.session_state.get("備考", "")
    }
    
    if save_meisai_as_json(st.session_state["見積No"], 保存データ):
        st.success("💾 明細データを保存しました")
        return True
    else:
        st.error("明細データの保存に失敗しました")
        return False

def export_estimate():
    """見積書を出力（係数対応版・明細番号修正版・エラーハンドリング強化）"""
    try:
        # 基本情報の準備
        見積No = st.session_state.get("見積No", "")
        案件名 = st.session_state.get("案件名", "案件名未設定")
        
        # 係数機能使用状況をチェック
        係数機能使用 = st.session_state.get("係数機能使用", False)
        
        # ファイル名に使用できない文字を置換（括弧はそのまま保持）
        import re
        安全な案件名 = 案件名
        
        # 危険な文字のみをアンダースコアに置換（括弧は問題ないのでそのまま）
        安全な案件名 = re.sub(r'[<>:"/\\|?*]', '_', 安全な案件名)
        
        # 前後の空白を除去
        安全な案件名 = 安全な案件名.strip()
        
        # ファイル名が長すぎる場合は短縮
        if len(安全な案件名) > 50:
            安全な案件名 = 安全な案件名[:50] + "..."
        
        ファイル名 = f"{見積No}見積書_{安全な案件名}.xlsx"
        
        # 住所情報の取得（細分化対応・優先順位修正版）
        郵便番号 = ""
        住所1 = ""
        住所2 = ""
        
        # 1. セッション状態から細分化住所を優先取得
        if st.session_state.get("選択された郵便番号") or st.session_state.get("選択された住所1") or st.session_state.get("選択された住所2"):
            郵便番号 = st.session_state.get("選択された郵便番号", "")
            住所1 = st.session_state.get("選択された住所1", "")
            住所2 = st.session_state.get("選択された住所2", "")
            st.info(f"セッション状態の細分化住所を使用: {郵便番号} {住所1} {住所2}")
        
        # 2. セッション状態に細分化住所がない場合は顧客JSONから取得
        else:
            try:
                顧客一覧 = load_customers_json()
                顧客会社名 = st.session_state.get("選択された顧客会社名", "")
                顧客担当者 = st.session_state.get("選択された顧客担当者", "")
                
                if 顧客会社名 and 顧客担当者 and 顧客一覧:
                    for customer in 顧客一覧:
                        if (customer.get("顧客会社名") == 顧客会社名 and 
                            customer.get("顧客担当者") == 顧客担当者):
                            
                            # 新形式の住所を優先
                            if customer.get("郵便番号") or customer.get("住所1") or customer.get("住所2"):
                                郵便番号 = customer.get("郵便番号", "")
                                住所1 = customer.get("住所1", "")
                                住所2 = customer.get("住所2", "")
                                st.info(f"顧客JSONの細分化住所を使用: {郵便番号} {住所1} {住所2}")
                            
                            # 新形式がない場合は旧住所を分割
                            elif customer.get("顧客住所"):
                                旧住所 = customer.get("顧客住所", "")
                                try:
                                    from estimate_excel_writer import parse_address
                                    郵便番号, 住所1, 住所2 = parse_address(旧住所)
                                    st.info(f"旧住所を分割して使用: {旧住所} → {郵便番号} {住所1} {住所2}")
                                except ImportError:
                                    st.warning("estimate_excel_writer モジュールが見つかりません")
                                    住所1 = 旧住所  # フォールバック
                                except Exception as e:
                                    st.warning(f"住所分割でエラー: {e}")
                                    住所1 = 旧住所  # フォールバック
                            
                            break
                
                if not (郵便番号 or 住所1 or 住所2):
                    st.warning("顧客JSONに住所データがありません")
                    
            except Exception as e:
                st.error(f"顧客住所の取得でエラーが発生しました: {e}")
        
        # 3. それでも住所が取得できない場合は統合住所から分割を試行
        if not (郵便番号 or 住所1 or 住所2):
            統合住所 = st.session_state.get("選択された顧客住所", "")
            if 統合住所:
                try:
                    from estimate_excel_writer import parse_address
                    郵便番号, 住所1, 住所2 = parse_address(統合住所)
                    st.info(f"統合住所を分割して使用: {統合住所} → {郵便番号} {住所1} {住所2}")
                except ImportError:
                    st.warning("estimate_excel_writer モジュールが見つかりません")
                    住所1 = 統合住所  # フォールバック
                except Exception as e:
                    st.warning(f"統合住所の分割でエラー: {e}")
                    住所1 = 統合住所  # フォールバック
        
        # デバッグ情報を表示
        st.write("**住所取得結果:**")
        st.write(f"- 郵便番号: '{郵便番号}'")
        st.write(f"- 住所1: '{住所1}'")
        st.write(f"- 住所2: '{住所2}'")
        
        if not (郵便番号 or 住所1 or 住所2):
            st.error("⚠️ 住所が設定されていません。見積書の住所欄は空になります。")
            st.info("💡 顧客情報の住所欄を確認して再度登録してください。")
        
        # 明細リストに商品番号を追加
        元明細リスト = st.session_state.get("明細リスト", [])
        処理済み明細リスト = []
        商品番号 = 0
        
        for item in 元明細リスト:
            処理済みitem = item.copy()
            
            # 分類項目の場合
            if item.get("分類", False):
                処理済みitem["商品番号"] = None  # 分類は番号なし
            else:
                # 商品項目の場合
                商品番号 += 1
                処理済みitem["商品番号"] = 商品番号
            
            処理済み明細リスト.append(処理済みitem)
        
        # 見積データを辞書として準備（係数機能フラグ追加）
        見積データ = {
            "見積No": 見積No,
            "案件名": 案件名,
            "発行日": st.session_state.get("発行日", datetime.date.today()),
            "顧客会社名": st.session_state.get("選択された顧客会社名", ""),
            "顧客部署名": st.session_state.get("選択された顧客部署名", ""),
            "顧客担当者": st.session_state.get("選択された顧客担当者", ""),
            "郵便番号": 郵便番号,
            "住所1": 住所1,
            "住所2": 住所2,
            "顧客住所": f"{郵便番号} {住所1} {住所2}".strip(),  # 互換性のため統合住所も設定
            "発行者名": st.session_state.get("発行者名", ""),
            "備考": st.session_state.get("備考", ""),
            "明細リスト": 処理済み明細リスト,  # 商品番号付きの明細リスト
            "係数機能使用": 係数機能使用  # 係数機能フラグを追加
        }
        
        # 見積書を生成
        try:
            from estimate_excel_writer import write_estimate_to_excel
            success = write_estimate_to_excel(見積データ, ファイル名)
        except ImportError:
            st.error("❌ estimate_excel_writer モジュールが見つかりません。")
            st.info("見積書出力機能を使用するには、estimate_excel_writer.py ファイルが必要です。")
            return
        except Exception as e:
            st.error(f"❌ 見積書生成でエラーが発生しました: {e}")
            return
        
        if success and os.path.exists(ファイル名):
            st.success(f"✅ 見積書を出力しました!")
            if 係数機能使用:
                st.info(f"📋 **係数対応テンプレートを使用しました**")
            else:
                st.info(f"📋 **通常テンプレートを使用しました**")
            st.info(f"📁 **保存先:** `{os.path.abspath(ファイル名)}`")
            
            # ダウンロードボタンを追加
            with open(ファイル名, "rb") as file:
                st.download_button(
                    label="📥 見積書をダウンロード",
                    data=file.read(),
                    file_name=ファイル名,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_estimate"
                )
        else:
            st.error("見積書の生成に失敗しました")
        
    except Exception as e:
        st.error(f"❌ エラーが発生しました: {e}")
        st.write("**デバッグ情報:**")
        st.write(f"- 見積No: {st.session_state.get('見積No', '未設定')}")
        st.write(f"- 明細件数: {len(st.session_state.get('明細リスト', []))}")
        st.write(f"- 通常テンプレート: {'存在' if os.path.exists('estimate_template.xlsx') else '存在しない'}")
        st.write(f"- 係数テンプレート: {'存在' if os.path.exists('estimate_templat_keisuu.xlsx') else '存在しない'}")
        st.write("**エラーの詳細:**")
        st.code(traceback.format_exc())

def render_editable_detail_list_with_coefficient(品名一覧_df):
    """編集可能な明細一覧を表示（係数対応版）"""
    st.subheader("📋 明細一覧（直接編集）")
    明細リスト = st.session_state["明細リスト"]
    係数機能使用 = st.session_state.get("係数機能使用", False)
    
    if not 明細リスト:
        st.info("明細が登録されていません。下記フォームから追加してください。")
        return
    
    # ヘッダー（係数機能に応じて変更）
    if 係数機能使用:
        header_col1, header_col2, header_col3, header_col4, header_col5, header_col6, header_col7, header_col8, header_col9, header_col10 = st.columns([0.5, 2, 1, 1, 1, 1.2, 1.2, 1.5, 1.5, 2])
        header_col1.write("**No**")
        header_col2.write("**品名**")
        header_col3.write("**数量**")
        header_col4.write("**単位**")
        header_col5.write("**係数**")
        header_col6.write("**単価**")
        header_col7.write("**金額**")
        header_col8.write("**売上先部署**")
        header_col9.write("**備考**")
        header_col10.write("**操作**")
    else:
        header_col1, header_col2, header_col3, header_col4, header_col5, header_col6, header_col7, header_col8, header_col9 = st.columns([0.5, 2, 1, 1, 1.2, 1.2, 1.5, 1.5, 2])
        header_col1.write("**No**")
        header_col2.write("**品名**")
        header_col3.write("**数量**")
        header_col4.write("**単位**")
        header_col5.write("**単価**")
        header_col6.write("**金額**")
        header_col7.write("**売上先部署**")
        header_col8.write("**備考**")
        header_col9.write("**操作**")
    
    # 商品番号カウンター（分類を除外）
    商品番号 = 0
    
    # 明細行（編集可能・分類対応・係数対応）
    for i, row in enumerate(明細リスト):
        編集中 = st.session_state.get(f"編集中_{i}", False)
        is_category = row.get("分類", False)  # 分類項目かどうか
        
        # 商品項目の場合のみ番号をカウントアップ
        if not is_category:
            商品番号 += 1
        
        if not 編集中:
            # 通常表示モード（係数機能対応）
            if 係数機能使用:
                col1, col2, col3, col4, col5, col6, col7, col8, col9, col10 = st.columns([0.5, 2, 1, 1, 1, 1.2, 1.2, 1.5, 1.5, 2])
            else:
                col1, col2, col3, col4, col5, col6, col7, col8, col9 = st.columns([0.5, 2, 1, 1, 1.2, 1.2, 1.5, 1.5, 2])
            
            # No列の表示：分類の場合は空、商品の場合は番号
            if is_category:
                col1.write("")  # 分類項目は番号なし
            else:
                col1.write(str(商品番号))  # 商品項目のみ番号表示
            
            # 分類項目の場合は特別な表示
            if is_category:
                col2.markdown(f"**{row['品名']}**")  # 太字で表示
                col3.write("-")
                col4.write("-")
                if 係数機能使用:
                    col5.write("-")
                    col6.write("-")
                    col7.write("-")
                    col8.write("-")
                    col9.write("-")
                else:
                    col5.write("-")
                    col6.write("-")
                    col7.write("-")
                    col8.write("-")
            else:
                # 通常の明細項目
                col2.write(row["品名"])
                
                # 数量の安全な表示
                try:
                    数量 = row.get("数量", 0)
                    if isinstance(数量, (int, float)):
                        col3.write(str(int(数量)))
                    else:
                        col3.write(str(数量))
                except:
                    col3.write("0")
                
                col4.write(row.get("単位", ""))
                
                if 係数機能使用:
                    # 係数の表示
                    係数 = row.get("係数", 1)
                    col5.write(str(係数))
                    
                    # 単価の安全な表示
                    try:
                        単価 = row.get("単価", 0)
                        if isinstance(単価, (int, float)):
                            col6.write(f"¥{単価:,}")
                        else:
                            col6.write(str(単価))
                    except:
                        col6.write("¥0")
                    
                    # 金額の安全な表示
                    try:
                        金額 = row.get("金額", 0)
                        if isinstance(金額, (int, float)):
                            col7.write(f"¥{金額:,}")
                        else:
                            col7.write(str(金額))
                    except:
                        col7.write("¥0")
                    
                    col8.write(row.get("売上先部署", ""))
                    col9.write(row.get("備考", ""))
                else:
                    # 単価の安全な表示
                    try:
                        単価 = row.get("単価", 0)
                        if isinstance(単価, (int, float)):
                            col5.write(f"¥{単価:,}")
                        else:
                            col5.write(str(単価))
                    except:
                        col5.write("¥0")
                    
                    # 金額の安全な表示
                    try:
                        金額 = row.get("金額", 0)
                        if isinstance(金額, (int, float)):
                            col6.write(f"¥{金額:,}")
                        else:
                            col6.write(str(金額))
                    except:
                        col6.write("¥0")
                    
                    col7.write(row.get("売上先部署", ""))
                    col8.write(row.get("備考", ""))
            
            # 操作ボタン（係数機能に関係なく最後の列）
            if 係数機能使用:
                操作col = col10
            else:
                操作col = col9
                
            with 操作col:
                button_col1, button_col2, button_col3, button_col4, button_col5 = st.columns(5)
                
                if button_col1.button("↑", key=f"up_{i}", help="上に移動", disabled=(i == 0)):
                    if i > 0:
                        明細リスト[i-1], 明細リスト[i] = 明細リスト[i], 明細リスト[i-1]
                        st.rerun()
                
                if button_col2.button("↓", key=f"down_{i}", help="下に移動", disabled=(i == len(明細リスト) - 1)):
                    if i < len(明細リスト) - 1:
                        明細リスト[i+1], 明細リスト[i] = 明細リスト[i], 明細リスト[i+1]
                        st.rerun()
                
                if button_col3.button("📝", key=f"edit_{i}", help="編集"):
                    # 他の編集をキャンセル
                    for j in range(len(明細リスト)):
                        st.session_state[f"編集中_{j}"] = False
                    st.session_state[f"編集中_{i}"] = True
                    st.rerun()
                
                if button_col4.button("📋", key=f"copy_{i}", help="複製"):
                    新明細 = row.copy()
                    明細リスト.insert(i + 1, 新明細)
                    st.success("明細を複製しました")
                    st.rerun()
                
                if button_col5.button("🗑️", key=f"delete_{i}", help="削除"):
                    明細リスト.pop(i)
                    st.success("明細を削除しました")
                    st.rerun()
        
        else:
            # 編集モード（係数対応版は長いため、次のartifactで提供）
            render_detail_edit_mode_with_coefficient(i, row, is_category, 商品番号)

    # 合計金額と部署別集計の表示（分類項目除外版）
    render_detail_summary()

def render_detail_edit_mode_with_coefficient(i, row, is_category, 商品番号):
    """明細編集モード（係数対応版）"""
    係数機能使用 = st.session_state.get("係数機能使用", False)
    明細リスト = st.session_state["明細リスト"]
    
    # 分類項目の編集
    if is_category:
        col1, col2 = st.columns([2, 1])
        
        with col1:
            品名 = st.text_input("分類名", value=row["品名"], key=f"edit_category_name_{i}")
        
        with col2:
            edit_col1, edit_col2 = st.columns(2)
            
            if edit_col1.button("✅ 保存", key=f"save_category_{i}"):
                明細リスト[i]["品名"] = 品名
                st.session_state[f"編集中_{i}"] = False
                st.success("分類を更新しました")
                st.rerun()
            
            if edit_col2.button("❌ キャンセル", key=f"cancel_category_{i}"):
                st.session_state[f"編集中_{i}"] = False
                st.rerun()
        
        return

    # 商品項目の編集（係数対応版）
    if 係数機能使用:
        col1, col2, col3, col4 = st.columns(4)
    else:
        col1, col2, col3 = st.columns(3)
    
    with col1:
        # 品名の編集
        品名 = st.text_input("品名", value=row["品名"], key=f"edit_name_{i}")
        
        # 数量の安全な編集
        try:
            現在の数量 = row.get("数量", 0)
            if 現在の数量 == "" or 現在の数量 is None:
                現在の数量 = 0
            else:
                現在の数量 = int(現在の数量)
        except (ValueError, TypeError):
            現在の数量 = 0
        
        数量 = st.number_input("数量", value=現在の数量, min_value=0, step=1, key=f"edit_qty_{i}")
        
        単位選択肢 = ["式", "分", "文字", "半日", "日", "時間", "名", "本", "件", "回", "枚", "個"]
        現在の単位 = row.get("単位", "式")
        単位初期index = 0
        if 現在の単位 in 単位選択肢:
            単位初期index = 単位選択肢.index(現在の単位)
        
        単位 = st.selectbox("単位", options=単位選択肢, index=単位初期index, key=f"edit_unit_{i}")
    
    with col2:
        if 係数機能使用:
            # 係数の安全な編集
            try:
                現在の係数 = row.get("係数", 1)
                if 現在の係数 == "" or 現在の係数 is None:
                    現在の係数 = 1.0
                else:
                    現在の係数 = float(現在の係数)
            except (ValueError, TypeError):
                現在の係数 = 1.0
            
            係数 = st.number_input("係数", value=現在の係数, min_value=0.5, step=0.5, key=f"edit_coeff_{i}")
        
        # 単価の安全な編集
        try:
            現在の単価 = row.get("単価", 0)
            if 現在の単価 == "" or 現在の単価 is None:
                現在の単価 = 0.0
            else:
                現在の単価 = float(現在の単価)
        except (ValueError, TypeError):
            現在の単価 = 0.0
        
        単価 = st.number_input("単価", value=現在の単価, step=100.0, help="割引の場合はマイナス値を入力してください", key=f"edit_price_{i}")
        
        # 金額計算
        if 係数機能使用:
            金額 = 数量 * 係数 * 単価
        else:
            金額 = 数量 * 単価
        
        st.write(f"**金額: ¥{金額:,}**")
    
    with col3:
        売上先部署選択肢 = ["", "映像制作部", "翻訳制作部", "完プロ制作部", "生字幕制作部", "字幕展開部"]
        現在の売上先部署 = row.get("売上先部署", "")
        売上先部署初期index = 0
        if 現在の売上先部署 in 売上先部署選択肢:
            売上先部署初期index = 売上先部署選択肢.index(現在の売上先部署)
        
        売上先部署 = st.selectbox("売上先部署", options=売上先部署選択肢, index=売上先部署初期index, key=f"edit_dept_{i}")
        
        備考 = st.text_input("備考", value=row.get("備考", ""), key=f"edit_remark_{i}")
    
    if 係数機能使用:
        with col4:
            st.write("")  # 高さ調整
            
            edit_col1, edit_col2 = st.columns(2)
            
            if edit_col1.button("✅ 保存", key=f"save_edit_{i}"):
                # 明細を更新
                明細リスト[i]["品名"] = 品名
                明細リスト[i]["数量"] = 数量
                明細リスト[i]["単位"] = 単位
                明細リスト[i]["係数"] = 係数
                明細リスト[i]["単価"] = 単価
                明細リスト[i]["金額"] = 金額
                明細リスト[i]["売上先部署"] = 売上先部署
                明細リスト[i]["備考"] = 備考
                
                st.session_state[f"編集中_{i}"] = False
                st.success("明細を更新しました")
                st.rerun()
            
            if edit_col2.button("❌ キャンセル", key=f"cancel_edit_{i}"):
                st.session_state[f"編集中_{i}"] = False
                st.rerun()
    else:
        # 係数機能未使用時の保存ボタン
        edit_col1, edit_col2 = st.columns([1, 1])
        
        if edit_col1.button("✅ 保存", key=f"save_edit_{i}"):
            # 明細を更新
            明細リスト[i]["品名"] = 品名
            明細リスト[i]["数量"] = 数量
            明細リスト[i]["単位"] = 単位
            明細リスト[i]["係数"] = 1  # 係数機能無効時は1で固定
            明細リスト[i]["単価"] = 単価
            明細リスト[i]["金額"] = 金額
            明細リスト[i]["売上先部署"] = 売上先部署
            明細リスト[i]["備考"] = 備考
            
            st.session_state[f"編集中_{i}"] = False
            st.success("明細を更新しました")
            st.rerun()
        
        if edit_col2.button("❌ キャンセル", key=f"cancel_edit_{i}"):
            st.session_state[f"編集中_{i}"] = False
            st.rerun()

def load_all_projects():
    """すべてのJSONファイルから案件データを読み込む（数値変換強化版）"""
    案件リスト = []
    
    if not os.path.exists(DATA_FOLDER):
        return []
    
    try:
        json_files = [f for f in os.listdir(DATA_FOLDER) if f.endswith('.json')]
        
        for file in json_files:
            try:
                ファイルパス = os.path.join(DATA_FOLDER, file)
                with open(ファイルパス, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                if isinstance(data, dict):
                    # 明細から合計金額を計算（分類項目除外・数値変換強化）
                    明細リスト = data.get("明細リスト", [])
                    明細合計 = 0
                    
                    for item in 明細リスト:
                        # 分類項目でない場合のみ合計に含める
                        if not item.get("分類", False):
                            try:
                                金額 = item.get("金額", 0)
                                # 空文字列や None の場合は 0 として扱う
                                if 金額 == "" or 金額 is None:
                                    金額 = 0
                                else:
                                    金額 = float(金額)
                                明細合計 += 金額
                            except (ValueError, TypeError):
                                # 変換できない場合は 0 として扱う
                                continue
                    
                    # 売上額の優先順位：明細合計 > 保存された売上額
                    try:
                        保存売上額 = data.get("売上額", 0)
                        if 保存売上額 == "" or 保存売上額 is None:
                            保存売上額 = 0
                        else:
                            保存売上額 = float(保存売上額)
                    except (ValueError, TypeError):
                        保存売上額 = 0
                    
                    # 明細がある場合は明細合計を優先、ない場合は保存された売上額を使用
                    if 明細合計 > 0:
                        売上額 = 明細合計
                    else:
                        売上額 = 保存売上額
                    
                    # 日付の処理
                    try:
                        発行日 = pd.to_datetime(data.get("発行日", "")).date() if data.get("発行日") else None
                    except:
                        発行日 = None
                    
                    try:
                        受注日 = pd.to_datetime(data.get("受注日", "")).date() if data.get("受注日") else None
                    except:
                        受注日 = None
                    
                    try:
                        納品日 = pd.to_datetime(data.get("納品日", "")).date() if data.get("納品日") else None
                    except:
                        納品日 = None
                    
                    # 数値フィールドの安全な変換
                    try:
                        仕入額 = data.get("仕入額", 0)
                        if 仕入額 == "" or 仕入額 is None:
                            仕入額 = 0
                        else:
                            仕入額 = int(float(仕入額))
                    except (ValueError, TypeError):
                        仕入額 = 0
                    
                    try:
                        粗利 = data.get("粗利", 0)
                        if 粗利 == "" or 粗利 is None:
                            粗利 = 0
                        else:
                            粗利 = int(float(粗利))
                    except (ValueError, TypeError):
                        粗利 = 0
                    
                    try:
                        粗利率 = data.get("粗利率", 0)
                        if 粗利率 == "" or 粗利率 is None:
                            粗利率 = 0
                        else:
                            粗利率 = float(粗利率)
                    except (ValueError, TypeError):
                        粗利率 = 0
                    
                    # 明細件数の計算（分類項目除外）
                    明細件数 = len([item for item in 明細リスト if not item.get("分類", False)])
                    
                    案件データ = {
                        "見積No": data.get("見積No", ""),
                        "案件名": data.get("案件名", "案件名未設定"),
                        "顧客会社名": data.get("顧客会社名", ""),
                        "顧客部署名": data.get("顧客部署名", ""),
                        "顧客担当者": data.get("顧客担当者", ""),
                        "発行日": 発行日,
                        "受注日": 受注日,
                        "納品日": 納品日,
                        "売上額": int(売上額),
                        "仕入額": 仕入額,
                        "粗利": 粗利,
                        "粗利率": 粗利率,
                        "状況": data.get("状況", "見積中"),
                        "発行者名": data.get("発行者名", ""),
                        "メモ": data.get("メモ", ""),
                        "明細件数": 明細件数,
                        "JSONファイル": file,
                        "担当部署": data.get("担当部署", ""),
                        "明細リスト": 明細リスト
                    }
                    
                    案件リスト.append(案件データ)
                    
            except Exception as e:
                st.warning(f"ファイル {file} の読み込みでエラー: {e}")
                continue
                
    except Exception as e:
        st.error(f"案件データの読み込みでエラー: {e}")
    
    return 案件リスト

def render_project_list_tab():
    """案件一覧タブを表示（年度・月連動フィルタ対応版）"""
    st.header("① 案件一覧")
    
    # 案件データの読み込み
    案件リスト = load_all_projects()
    
    if not 案件リスト:
        st.info("案件データがありません。")
        if st.button("新しい案件を作成", key="create_project_no_data"):
            # データを自動リセット
            reset_all_data()
            st.session_state["アクティブタブ"] = "② 顧客情報を入力"
            st.success("新しい案件を作成します。データをリセットしました。")
            st.rerun()
    
    # ヘッダー：フィルタ・検索　「絞り込む」ボタン　「すべてクリア」ボタン
    header_col1, header_col2, header_col3 = st.columns([2, 1, 1])
    with header_col1:
        st.subheader("🔍 フィルタ・検索")
    with header_col2:
        st.write("　")  # 高さ調整
        絞り込み実行 = st.button("🔍 絞り込む", type="primary", key="apply_filter")
    with header_col3:
        st.write("　")  # 高さ調整
        if st.button("🧹 すべてクリア", key="clear_all_filters", help="すべてのフィルタと検索をクリア"):
            # チェックボックス関連のキーを削除
            keys_to_remove = []
            for key in list(st.session_state.keys()):
                if key.startswith("include_") or key.startswith("exclude_"):
                    keys_to_remove.append(key)
            
            for key in keys_to_remove:
                if key in st.session_state:
                    del st.session_state[key]
            
            # フィルタ値を初期化
            st.session_state["filter_売上年度"] = "すべて"
            st.session_state["filter_売上月"] = "すべて"
            st.session_state["filter_顧客"] = "すべて"
            st.session_state["filter_発行者"] = "すべて"
            st.session_state["filter_担当部署"] = "すべて"
            st.session_state["filter_検索キーワード"] = ""
            
            st.rerun()

    # 1行目：案件名で検索
    検索キーワード = st.text_input(
        "🔍 案件名で検索", 
        placeholder="案件名を入力してください",
        value=st.session_state.get("filter_検索キーワード", ""),
        key="search_input"
    )
    st.session_state["filter_検索キーワード"] = 検索キーワード
    
    # 2行目：売上年度、売上月、顧客、発行者、担当部署（5等分）
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        # 売上年度でフィルタ（年度計算修正版）
        売上年度リスト = ["すべて"]
        年度セット = set()
        
        # 案件データを読み込んで年度を抽出
        案件リスト = load_all_projects()
        
        # 納品日から年度を抽出（4月-3月ベース）
        for 案件 in 案件リスト:
            if 案件["納品日"]:
                try:
                    納品日 = 案件["納品日"]
                    # 年度計算（4月-3月）
                    if 納品日.month >= 4:
                        年度 = 納品日.year
                    else:
                        年度 = 納品日.year - 1
                    年度セット.add(年度)
                except:
                    pass
        
        # 年度を降順でソートしてリストに追加
        if 年度セット:
            年度リスト = sorted(list(年度セット), reverse=True)
            for 年度 in 年度リスト:
                売上年度リスト.append(f"{年度}年度")
        
        # 初期値を設定
        初期年度index = 0
        if st.session_state.get("filter_売上年度") in 売上年度リスト:
            初期年度index = 売上年度リスト.index(st.session_state.get("filter_売上年度"))
        
        選択年度 = st.selectbox(
            "売上年度", 
            売上年度リスト,
            index=初期年度index,
            key="select_年度"
        )
        st.session_state["filter_売上年度"] = 選択年度
    
    with col2:
        # 売上月でフィルタ（月のみの選択に変更）
        売上月リスト = ["すべて", "1月", "2月", "3月", "4月", "5月", "6月", 
                        "7月", "8月", "9月", "10月", "11月", "12月"]
        
        # 初期値を設定
        初期月index = 0
        if st.session_state.get("filter_売上月") in 売上月リスト:
            初期月index = 売上月リスト.index(st.session_state.get("filter_売上月"))
        
        選択月 = st.selectbox(
            "売上月", 
            売上月リスト,
            index=初期月index,
            key="select_月"
        )
        st.session_state["filter_売上月"] = 選択月

    with col3:
        # 顧客でフィルタ（顧客No順で並び替え）
        try:
            # 顧客JSONから顧客No順で会社名を取得
            顧客一覧 = load_customers_json()
            if 顧客一覧:
                # 顧客Noでソートしてから会社名を取得
                顧客一覧.sort(key=lambda x: (x.get("顧客No", 999), x.get("顧客会社名", "")))
                顧客会社名リスト = []
                for 顧客 in 顧客一覧:
                    会社名 = 顧客.get("顧客会社名", "")
                    if 会社名 and 会社名 not in 顧客会社名リスト:
                        顧客会社名リスト.append(会社名)
                
                顧客リスト = ["すべて"] + 顧客会社名リスト
            else:
                # 顧客JSONが空の場合は案件データから取得（フォールバック）
                顧客リスト = ["すべて"] + sorted(list(set([案件["顧客会社名"] for 案件 in 案件リスト if 案件["顧客会社名"]])))
        except:
            # エラーの場合は従来通り
            顧客リスト = ["すべて"] + sorted(list(set([案件["顧客会社名"] for 案件 in 案件リスト if 案件["顧客会社名"]])))
        
        # 初期値を設定
        初期顧客index = 0
        if st.session_state.get("filter_顧客") in 顧客リスト:
            初期顧客index = 顧客リスト.index(st.session_state.get("filter_顧客"))
        
        選択顧客 = st.selectbox(
            "顧客", 
            顧客リスト,
            index=初期顧客index,
            key="select_顧客"
        )
        st.session_state["filter_顧客"] = 選択顧客
    
    with col4:
        # 発行者でフィルタ（既存コード）
        発行者リスト = ["すべて"] + sorted(list(set([案件["発行者名"] for 案件 in 案件リスト if 案件["発行者名"]])))
        
        # 初期値を設定
        初期発行者index = 0
        if st.session_state.get("filter_発行者") in 発行者リスト:
            初期発行者index = 発行者リスト.index(st.session_state.get("filter_発行者"))
        
        選択発行者 = st.selectbox(
            "発行者", 
            発行者リスト,
            index=初期発行者index,
            key="select_発行者"
        )
        st.session_state["filter_発行者"] = 選択発行者
    
    with col5:
        # 担当部署でフィルタ（新規追加）
        担当部署リスト = ["すべて"] + sorted(list(set([案件["担当部署"] for 案件 in 案件リスト if 案件.get("担当部署")])))
        
        # 初期値を設定
        初期担当部署index = 0
        if st.session_state.get("filter_担当部署") in 担当部署リスト:
            初期担当部署index = 担当部署リスト.index(st.session_state.get("filter_担当部署"))
        
        選択担当部署 = st.selectbox(
            "担当部署", 
            担当部署リスト,
            index=初期担当部署index,
            key="select_担当部署"
        )
        st.session_state["filter_担当部署"] = 選択担当部署

    # 4行目：状況（含む）、状況（除く）ヘッダー（2等分・背景色付き）
    col1, col2 = st.columns(2)
    
    with col1:
        # 状況（含む）のヘッダーに青い背景色
        st.info("**状況（含む）**")
    
    with col2:
        # 状況（除く）のヘッダーに黄色い背景色
        st.warning("**状況（除く）**")
    
    # 5行目：クリアボタン（2等分）
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🧹 含むをクリア", key="clear_include_status", help="状況（含む）の選択をすべて解除"):
            # 含むチェックボックスをすべてクリア
            for 状況 in ["見積中", "受注", "納品済", "請求済", "不採用", "失注"]:
                st.session_state[f"include_{状況}"] = False
            st.rerun()
    
    with col2:
        if st.button("🧹 除くをクリア", key="clear_exclude_status", help="状況（除く）の選択をすべて解除"):
            # 除くチェックボックスをすべてクリア
            for 状況 in ["見積中", "受注", "納品済", "請求済", "不採用", "失注"]:
                st.session_state[f"exclude_{状況}"] = False
            st.rerun()
    
    # 6行目以降：チェックボックス（2等分、シンプル表示）
    col1, col2 = st.columns(2)
    
    状況選択肢 = ["見積中", "受注", "納品済", "請求済", "不採用", "失注"]
    
    with col1:
        # 状況（含む）のチェックボックス - シンプル表示
        # 1段目：見積中、受注、納品済
        check_row1_col1, check_row1_col2, check_row1_col3 = st.columns(3)
        with check_row1_col1:
            見積中_checked = st.checkbox("見積中", key="include_見積中_ui", value=st.session_state.get("include_見積中", False))
            st.session_state["include_見積中"] = 見積中_checked
        with check_row1_col2:
            受注_checked = st.checkbox("受注", key="include_受注_ui", value=st.session_state.get("include_受注", False))
            st.session_state["include_受注"] = 受注_checked
        with check_row1_col3:
            納品済_checked = st.checkbox("納品済", key="include_納品済_ui", value=st.session_state.get("include_納品済", False))
            st.session_state["include_納品済"] = 納品済_checked
        
        # 2段目：請求済、不採用、失注
        check_row2_col1, check_row2_col2, check_row2_col3 = st.columns(3)
        with check_row2_col1:
            請求済_checked = st.checkbox("請求済", key="include_請求済_ui", value=st.session_state.get("include_請求済", False))
            st.session_state["include_請求済"] = 請求済_checked
        with check_row2_col2:
            不採用_checked = st.checkbox("不採用", key="include_不採用_ui", value=st.session_state.get("include_不採用", False))
            st.session_state["include_不採用"] = 不採用_checked
        with check_row2_col3:
            失注_checked = st.checkbox("失注", key="include_失注_ui", value=st.session_state.get("include_失注", False))
            st.session_state["include_失注"] = 失注_checked
    
    with col2:
        # 状況（除く）のチェックボックス - シンプル表示
        # 1段目：見積中、受注、納品済
        exclude_row1_col1, exclude_row1_col2, exclude_row1_col3 = st.columns(3)
        with exclude_row1_col1:
            見積中_excluded = st.checkbox("見積中", key="exclude_見積中_ui", value=st.session_state.get("exclude_見積中", False))
            st.session_state["exclude_見積中"] = 見積中_excluded
        with exclude_row1_col2:
            受注_excluded = st.checkbox("受注", key="exclude_受注_ui", value=st.session_state.get("exclude_受注", False))
            st.session_state["exclude_受注"] = 受注_excluded
        with exclude_row1_col3:
            納品済_excluded = st.checkbox("納品済", key="exclude_納品済_ui", value=st.session_state.get("exclude_納品済", False))
            st.session_state["exclude_納品済"] = 納品済_excluded
        
        # 2段目：請求済、不採用、失注
        exclude_row2_col1, exclude_row2_col2, exclude_row2_col3 = st.columns(3)
        with exclude_row2_col1:
            請求済_excluded = st.checkbox("請求済", key="exclude_請求済_ui", value=st.session_state.get("exclude_請求済", False))
            st.session_state["exclude_請求済"] = 請求済_excluded
        with exclude_row2_col2:
            不採用_excluded = st.checkbox("不採用", key="exclude_不採用_ui", value=st.session_state.get("exclude_不採用", False))
            st.session_state["exclude_不採用"] = 不採用_excluded
        with exclude_row2_col3:
            失注_excluded = st.checkbox("失注", key="exclude_失注_ui", value=st.session_state.get("exclude_失注", False))
            st.session_state["exclude_失注"] = 失注_excluded

    # 絞り込みボタンが押された場合のみフィルタを適用
    if 絞り込み実行:
        選択された売上年度 = st.session_state.get("filter_売上年度", "すべて")
        選択された売上月 = st.session_state.get("filter_売上月", "すべて")
        選択された顧客 = st.session_state.get("filter_顧客", "すべて")
        選択された発行者 = st.session_state.get("filter_発行者", "すべて")
        選択された担当部署 = st.session_state.get("filter_担当部署", "すべて")
        検索キーワード = st.session_state.get("filter_検索キーワード", "")
        
        # 状況フィルタを取得（修正された状況リストを使用）
        選択された状況含む = []
        選択された状況除く = []
        for 状況 in ["見積中", "受注", "納品済", "請求済", "不採用", "失注"]:
            if st.session_state.get(f"include_{状況}", False):
                選択された状況含む.append(状況)
            if st.session_state.get(f"exclude_{状況}", False):
                選択された状況除く.append(状況)
    else:
        # 絞り込みボタンが押されていない場合は全件表示
        選択された売上年度 = "すべて"
        選択された売上月 = "すべて"
        選択された顧客 = "すべて"
        選択された発行者 = "すべて"
        選択された担当部署 = "すべて"
        検索キーワード = ""
        選択された状況含む = []
        選択された状況除く = []

    # フィルタ適用処理（年度・月連動対応版）
    フィルタ済み案件 = 案件リスト.copy()
    
    # 年度・月の連動フィルタ
    if 選択された売上年度 != "すべて" or 選択された売上月 != "すべて":
        フィルタ済み案件 = []
        for 案件 in 案件リスト:
            if not 案件["納品日"]:
                continue
            
            納品日 = 案件["納品日"]
            
            # 年度の計算（4月-3月ベース）
            if 納品日.month >= 4:
                案件年度 = 納品日.year
            else:
                案件年度 = 納品日.year - 1
            
            # 年度フィルタのチェック
            年度一致 = True
            if 選択された売上年度 != "すべて":
                指定年度 = int(選択された売上年度.replace("年度", ""))
                年度一致 = (案件年度 == 指定年度)
            
            # 月フィルタのチェック
            月一致 = True
            if 選択された売上月 != "すべて":
                指定月 = int(選択された売上月.replace("月", ""))
                月一致 = (納品日.month == 指定月)
            
            # 両方の条件を満たす場合のみ追加
            if 年度一致 and 月一致:
                フィルタ済み案件.append(案件)
    
    if 選択された発行者 != "すべて":
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 案件["発行者名"] == 選択された発行者]
    
    if 選択された顧客 != "すべて":
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 案件["顧客会社名"] == 選択された顧客]
    
    # 担当部署でのフィルタ（新規追加）
    if 選択された担当部署 != "すべて":
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 案件.get("担当部署") == 選択された担当部署]
    
    # 状況（含む）フィルタ - 複数選択対応
    if 選択された状況含む:  # リストが空でない場合のみ適用
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 案件["状況"] in 選択された状況含む]
    
    # 状況（除く）フィルタ - 複数選択対応
    if 選択された状況除く:  # リストが空でない場合のみ適用
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 案件["状況"] not in 選択された状況除く]
    
    if 検索キーワード:
        フィルタ済み案件 = [案件 for 案件 in フィルタ済み案件 if 検索キーワード in 案件["案件名"]]
    
    # 統計情報
    st.subheader("📊 統計情報")
    
    # フィルタ済み案件での統計計算
    フィルタ済み件数 = len(フィルタ済み案件)
    
    # 売上見込み（受注～請求済み）
    売上見込み状況 = ["受注", "納品済", "請求済"]
    売上見込み = sum([
        案件["売上額"] for 案件 in フィルタ済み案件 
        if 案件["状況"] in 売上見込み状況
    ])
    
    # 売上合計（請求済みのみ）
    売上合計 = sum([
        案件["売上額"] for 案件 in フィルタ済み案件 
        if 案件["状況"] == "請求済"
    ])
    
    # 見積中件数
    見積中件数 = len([案件 for 案件 in フィルタ済み案件 if 案件["状況"] == "見積中"])
    
    # 請求済み件数（新規追加）
    請求済み件数 = len([案件 for 案件 in フィルタ済み案件 if 案件["状況"] == "請求済"])
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("案件数", フィルタ済み件数)
    
    with col2:
        st.metric("見積中", 見積中件数)
    
    with col3:
        st.metric("請求済み", 請求済み件数)
    
    with col4:
        st.metric("売上見込み", f"¥{売上見込み:,}")
    
    with col5:
        st.metric("売上合計", f"¥{売上合計:,}")
    
    # 選択された条件の表示（年度・月連動対応）
    表示メッセージ = []
    if 選択された売上年度 != "すべて" and 選択された売上月 != "すべて":
        # 年度と月が両方選択されている場合
        年度 = int(選択された売上年度.replace("年度", ""))
        月 = int(選択された売上月.replace("月", ""))
        
        # 年度に基づいて実際の年を計算
        if 月 >= 4:
            実年 = 年度
        else:
            実年 = 年度 + 1
        
        表示メッセージ.append(f"📅 {実年}年{月}月（{選択された売上年度}）")
    
    elif 選択された売上年度 != "すべて":
        表示メッセージ.append(f"📅 {選択された売上年度}")
    elif 選択された売上月 != "すべて":
        表示メッセージ.append(f"📅 {選択された売上月}")
    
    if 選択された状況含む:
        表示メッセージ.append(f"🎯 含む: {', '.join(選択された状況含む)}")
    if 選択された状況除く:
        表示メッセージ.append(f"❌ 除く: {', '.join(選択された状況除く)}")
    
    if 表示メッセージ:
        st.info(" | ".join(表示メッセージ) + " の条件で表示中")
    
    st.divider()
    
    # 案件一覧の表示（表示改善版）
    st.subheader(f"📋 案件一覧 ({len(フィルタ済み案件)}件)")

    if not フィルタ済み案件:
        st.info("条件に合致する案件がありません。")
        return

    # 案件一覧を表示（文字サイズ拡大・部署別集計改善）
    for i, 案件 in enumerate(フィルタ済み案件):
        # 部署別集計を計算（改善版）
        部署別集計表示 = ""
        総額 = 案件.get("売上額", 0)
        
        try:
            明細リスト = 案件.get("明細リスト", [])
            案件担当部署 = 案件.get("担当部署", "")
            
            if 明細リスト:
                部署別集計 = {}
                明細合計 = 0
                
                for row in 明細リスト:
                    金額 = row.get("金額", 0)
                    明細合計 += 金額
                    
                    売上先部署 = row.get("売上先部署", "")
                    使用部署 = 売上先部署 if 売上先部署 else 案件担当部署
                    
                    if 使用部署:
                        if 使用部署 in 部署別集計:
                            部署別集計[使用部署] += 金額
                        else:
                            部署別集計[使用部署] = 金額
                
                # 部署別集計の表示形式を改善
                if 部署別集計:
                    部署数 = len(部署別集計)
                    if 部署数 > 1:
                        # 複数部署の場合：総額を先頭に表示
                        部署別表示リスト = [f"総額:¥{総額:,}"]
                        for 部署, 金額 in sorted(部署別集計.items()):
                            部署別表示リスト.append(f"{部署}:¥{金額:,}")
                        部署別集計表示 = f" ({' | '.join(部署別表示リスト)})"
                    else:
                        # 単一部署の場合：従来通り
                        部署, 金額 = list(部署別集計.items())[0]
                        部署別集計表示 = f" ({部署}:¥{金額:,})"
        except:
            pass

        # 案件タイトルに部署別集計を含める（文字サイズ拡大）
        タイトル = f"📄 {案件['見積No']} - {案件['案件名']} ({案件['状況']}){部署別集計表示}"

        # Markdownで文字サイズを拡大
        with st.expander(タイトル):
            # expanderのタイトル部分を大きく表示するためのCSS適用
            st.markdown(f"""
            <style>
            .streamlit-expanderHeader {{
                font-size: 18px !important;
                font-weight: bold !important;
            }}
            </style>
            """, unsafe_allow_html=True)
            
            # 基本情報
            col1, col2 = st.columns(2)
            
            with col1:
                st.write(f"**見積No:** {案件['見積No']}")
                st.write(f"**案件名:** {案件['案件名']}")
                st.write(f"**顧客会社名:** {案件['顧客会社名']}")
                st.write(f"**顧客担当者:** {案件['顧客担当者']}")
                st.write(f"**発行者:** {案件['発行者名']}")
            
            with col2:
                st.write(f"**状況:** {案件['状況']}")
                st.write(f"**発行日:** {案件['発行日'] or '未設定'}")
                st.write(f"**受注日:** {案件['受注日'] or '未設定'}")
                st.write(f"**納品日:** {案件['納品日'] or '未設定'}")
            
            # 金額情報
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("売上額", f"¥{案件['売上額']:,}")
            with col2:
                st.metric("仕入額", f"¥{案件['仕入額']:,}")
            with col3:
                if 案件['粗利'] >= 0:
                    st.metric("粗利", f"¥{案件['粗利']:,}")
                else:
                    st.metric("粗利", f"¥{案件['粗利']:,}", delta_color="inverse")
            with col4:
                st.metric("粗利率", f"{案件['粗利率']:.1f}%")
            
            # メモ
            if 案件['メモ']:
                st.write(f"**メモ:** {案件['メモ']}")
            
            # 明細の部署別集計を表示
            try:
                ファイルパス = os.path.join(DATA_FOLDER, 案件['JSONファイル'])
                with open(ファイルパス, "r", encoding="utf-8") as f:
                    詳細データ = json.load(f)
                
                if isinstance(詳細データ, dict):
                    明細リスト = 詳細データ.get("明細リスト", [])
                    案件担当部署 = 詳細データ.get("担当部署", "")
                    
                    if 明細リスト:
                        # 部署別集計の計算
                        部署別集計 = {}
                        明細合計 = 0
                        
                        for row in 明細リスト:
                            金額 = row.get("金額", 0)
                            明細合計 += 金額
                            
                            売上先部署 = row.get("売上先部署", "")
                            # 売上先部署が未設定の場合は案件の担当部署を使用
                            使用部署 = 売上先部署 if 売上先部署 else 案件担当部署
                            
                            if 使用部署:
                                if 使用部署 in 部署別集計:
                                    部署別集計[使用部署] += 金額
                                else:
                                    部署別集計[使用部署] = 金額
                        
                        # 明細合計の表示
                        st.write(f"**明細合計:** ¥{明細合計:,}")
                        
                        # 部署別集計の表示
                        if 部署別集計:
                            部署別表示リスト = []
                            for 部署, 金額 in sorted(部署別集計.items()):
                                部署別表示リスト.append(f"{部署}: ¥{金額:,}")
                            
                            if 部署別表示リスト:
                                st.write(f"**部署別内訳:** {' | '.join(部署別表示リスト)}")
                        
                        st.write(f"**明細件数:** {len(明細リスト)}件")
            except Exception as e:
                pass  # エラーが発生しても案件表示は継続
            
            # 操作ボタン
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                if st.button("📝 編集", key=f"edit_{案件['見積No']}"):
                    # 案件データを読み込んで編集モードに（タブ③に移動）
                    if auto_load_json_by_estimate_no(案件['見積No'], auto_rerun=False):
                        st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                        st.success(f"案件 {案件['見積No']} を編集モードで開きました")
                        st.rerun()
                    else:
                        st.error("案件データの読み込みに失敗しました")
            
            with col2:
                if st.button("📋 明細表示", key=f"detail_{案件['見積No']}"):
                    # 案件データを読み込んで明細表示（タブ④に移動）
                    if auto_load_json_by_estimate_no(案件['見積No'], auto_rerun=False):
                        st.session_state["アクティブタブ"] = "④ 明細情報を入力"
                        st.success(f"案件 {案件['見積No']} の明細を表示しました")
                        st.rerun()
                    else:
                        st.error("明細データの読み込みに失敗しました")
            
            with col3:
                if st.button("📄 コピー", key=f"copy_{案件['見積No']}"):
                    # 案件をコピーして新規案件作成モードに
                    if copy_project_data(案件['見積No']):
                        st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                        st.success(f"案件 {案件['見積No']} をコピーしました。日付を更新して保存してください。")
                        st.rerun()
                    else:
                        st.error("案件のコピーに失敗しました")
            
            with col4:
                # 削除確認状態をチェック
                削除確認中 = st.session_state.get(f"削除確認_{案件['見積No']}", False)
                
                if not 削除確認中:
                    # 通常の削除ボタン
                    if st.button("🗑️ 削除", key=f"delete_{案件['見積No']}"):
                        st.session_state[f"削除確認_{案件['見積No']}"] = True
                        st.rerun()
                else:
                    # 削除確認中の表示
                    st.warning("本当に削除しますか？")
            
            # 削除確認ダイアログ（削除確認中の場合のみ表示）
            if st.session_state.get(f"削除確認_{案件['見積No']}", False):
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.button("✅ はい、削除します", key=f"confirm_delete_{案件['見積No']}", type="primary"):
                        # JSONファイルを削除
                        try:
                            ファイルパス = os.path.join(DATA_FOLDER, 案件['JSONファイル'])
                            if os.path.exists(ファイルパス):
                                os.remove(ファイルパス)
                                st.success(f"案件 {案件['見積No']} を削除しました")
                                # 削除確認フラグをクリア
                                del st.session_state[f"削除確認_{案件['見積No']}"]
                                st.rerun()
                            else:
                                st.error("削除対象のファイルが見つかりません")
                                # エラーでも確認フラグはクリア
                                del st.session_state[f"削除確認_{案件['見積No']}"]
                        except Exception as e:
                            st.error(f"削除に失敗しました: {e}")
                            # エラーでも確認フラグはクリア
                            del st.session_state[f"削除確認_{案件['見積No']}"]
                
                with col2:
                    if st.button("❌ キャンセル", key=f"cancel_delete_{案件['見積No']}"):
                        # 削除確認フラグをクリア
                        del st.session_state[f"削除確認_{案件['見積No']}"]
                        st.rerun()

    # 新規案件作成ボタン
    st.divider()
    if st.button("➕ 新しい案件を作成", type="primary", key="create_new_project"):
        # データを自動リセット
        reset_all_data()
        st.session_state["アクティブタブ"] = "② 顧客情報を入力"
        st.success("新しい案件を作成します。データをリセットしました。")
        st.rerun()

# コピー機能の関数（新規追加）
def copy_project_data(元見積No):
    """案件データをコピーして新規案件として設定（住所引き継ぎ強化版・エラー修正）"""
    try:
        ファイルパス = os.path.join(DATA_FOLDER, f"{元見積No}.json")
        
        if not os.path.exists(ファイルパス):
            return False

        with open(ファイルパス, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        if not isinstance(data, dict):
            return False

        # 新しい見積番号を生成
        今日 = datetime.date.today()
        新見積No = generate_estimate_no(今日)
        
        # 元データから顧客住所を確実に取得
        元顧客住所 = data.get("顧客住所", "")
        元郵便番号 = data.get("郵便番号", "")
        元住所1 = data.get("住所1", "")
        元住所2 = data.get("住所2", "")
        
        # 細分化住所の優先取得
        if 元郵便番号 or 元住所1 or 元住所2:
            # 新形式の住所がある場合はそのまま使用
            pass
        elif 元顧客住所 and str(元顧客住所) != "nan":
            # 旧形式の住所のみある場合は分割を試行
            try:
                from estimate_excel_writer import parse_address
                元郵便番号, 元住所1, 元住所2 = parse_address(元顧客住所)
            except ImportError:
                st.warning("estimate_excel_writer モジュールが見つかりません")
                元住所1 = 元顧客住所  # フォールバック
            except Exception:
                元住所1 = 元顧客住所  # フォールバック
        else:
            # JSONに住所がない場合は顧客一覧から取得を試行
            try:
                顧客一覧 = load_customers_json()
                元顧客会社名 = data.get("顧客会社名", "")
                元顧客担当者 = data.get("顧客担当者", "")
                
                if 元顧客会社名 and 元顧客担当者 and 顧客一覧:
                    for customer in 顧客一覧:
                        if (customer.get("顧客会社名") == 元顧客会社名 and 
                            customer.get("顧客担当者") == 元顧客担当者):
                            # 新形式優先
                            if customer.get("郵便番号") or customer.get("住所1") or customer.get("住所2"):
                                元郵便番号 = customer.get("郵便番号", "")
                                元住所1 = customer.get("住所1", "")
                                元住所2 = customer.get("住所2", "")
                            else:
                                元顧客住所 = customer.get("顧客住所", "")
                                if 元顧客住所:
                                    try:
                                        from estimate_excel_writer import parse_address
                                        元郵便番号, 元住所1, 元住所2 = parse_address(元顧客住所)
                                    except ImportError:
                                        元住所1 = 元顧客住所
                                    except Exception:
                                        元住所1 = 元顧客住所
                            break
            except Exception:
                pass
        
        # データをセッション状態にコピー（日付は今日に設定）
        st.session_state["見積No"] = 新見積No
        st.session_state["案件名"] = data.get("案件名", "") + "（コピー）"
        st.session_state["発行日"] = 今日  # 今日の日付に変更
        st.session_state["選択された顧客会社名"] = data.get("顧客会社名", "")
        st.session_state["選択された顧客部署名"] = data.get("顧客部署名", "")
        st.session_state["選択された顧客担当者"] = data.get("顧客担当者", "")
        
        # 住所情報の設定（細分化対応）
        st.session_state["選択された郵便番号"] = 元郵便番号
        st.session_state["選択された住所1"] = 元住所1
        st.session_state["選択された住所2"] = 元住所2
        統合住所 = f"{元郵便番号} {元住所1} {元住所2}".strip()
        st.session_state["選択された顧客住所"] = 統合住所 if 統合住所 else 元顧客住所
        
        st.session_state["発行者名"] = data.get("発行者名", ISSUER_LIST[0])
        st.session_state["担当部署"] = data.get("担当部署", "")
        st.session_state["明細リスト"] = data.get("明細リスト", [])
        st.session_state["備考"] = data.get("備考", "")
        st.session_state["売上額"] = int(data.get("売上額", 0))
        st.session_state["仕入額"] = int(data.get("仕入額", 0))
        st.session_state["粗利"] = int(data.get("粗利", 0))
        st.session_state["粗利率"] = float(data.get("粗利率", 0.0))
        st.session_state["状況"] = "見積中"  # 新規案件なので見積中に設定
        st.session_state["受注日"] = None    # 日付は空に設定
        st.session_state["納品日"] = None    # 日付は空に設定
        st.session_state["メモ"] = data.get("メモ", "")
        
        # デバッグ情報
        if 統合住所 or 元顧客住所:
            st.info(f"住所情報をコピーしました: {統合住所 or 元顧客住所}")
        else:
            st.warning("元案件に住所情報がありません。顧客情報で住所を確認してください。")
        
        return True
        
    except Exception as e:
        st.error(f"コピー処理でエラーが発生しました: {e}")
        return False

# 以下の関数を既存のコードに追加してください

def move_single_product(products_list, index, direction):
    """単一商品の位置を移動（JSONに保存）"""
    try:
        if direction == "up" and index > 0:
            # 上に移動
            products_list[index-1], products_list[index] = products_list[index], products_list[index-1]
        elif direction == "down" and index < len(products_list) - 1:
            # 下に移動
            products_list[index+1], products_list[index] = products_list[index], products_list[index+1]
        
        # JSONに保存
        if save_products_json(products_list):
            st.success(f"商品の並び順を変更しました")
        else:
            st.error("並び順の保存に失敗しました")
            
    except Exception as e:
        st.error(f"並び替えでエラーが発生しました: {e}")

def move_selected_products(products_list, direction):
    """選択された商品を一括移動（JSONに保存）"""
    try:
        selected_names = st.session_state.get("selected_products", set())
        if not selected_names:
            st.warning("移動する商品を選択してください")
            return
        
        # 選択された商品のインデックスを取得
        selected_indices = []
        for i, product in enumerate(products_list):
            if product.get("品名") in selected_names:
                selected_indices.append(i)
        
        if not selected_indices:
            st.warning("選択された商品が見つかりません")
            return
        
        # 移動処理
        if direction == "up":
            # 上に移動（前から処理）
            selected_indices.sort()
            for i in selected_indices:
                if i > 0:
                    # 移動先が選択済み商品でない場合のみ移動
                    target_index = i - 1
                    if target_index not in selected_indices:
                        products_list[target_index], products_list[i] = products_list[i], products_list[target_index]
                        # インデックスを更新
                        selected_indices = [idx-1 if idx == i else idx for idx in selected_indices]
        
        elif direction == "down":
            # 下に移動（後ろから処理）
            selected_indices.sort(reverse=True)
            for i in selected_indices:
                if i < len(products_list) - 1:
                    # 移動先が選択済み商品でない場合のみ移動
                    target_index = i + 1
                    if target_index not in selected_indices:
                        products_list[target_index], products_list[i] = products_list[i], products_list[target_index]
                        # インデックスを更新
                        selected_indices = [idx+1 if idx == i else idx for idx in selected_indices]
        
        # JSONに保存
        if save_products_json(products_list):
            st.success(f"選択した{len(selected_names)}件の商品を{direction}に移動しました")
        else:
            st.error("並び順の保存に失敗しました")
            
    except Exception as e:
        st.error(f"一括移動でエラーが発生しました: {e}")

# load_excel_data関数も必要ですが、使用していないようなので削除するか、以下のダミー関数を追加
def load_excel_data(filename):
    """Excelデータ読み込み（ダミー関数）"""
    # JSONのみを使用するため、空のDataFrameを返す
    import pandas as pd
    顧客一覧 = pd.DataFrame(columns=["顧客No", "顧客会社名", "顧客部署名", "顧客担当者", "顧客住所"])
    案件一覧 = pd.DataFrame(columns=["見積No", "案件名", "顧客会社名", "顧客部署名", "顧客担当者", "発行日", "受注日", "納品日", "売上額", "仕入額", "粗利", "粗利率", "状況", "発行者名", "メモ"])
    品名一覧 = pd.DataFrame(columns=["品名", "単位", "単価", "備考"])
    return 顧客一覧, 案件一覧, 品名一覧

def render_product_list_tab():
    """商品一覧タブを表示（一括削除・言語統合対応版）"""
    st.header("⑥ 商品一覧")

    # 新規商品追加ボタン（ヘッダー直下に配置）
    if st.button("➕ 新しい商品を追加", type="primary", key="add_product_header"):
        st.session_state["商品追加モード"] = True
        st.rerun()

    # 商品データの読み込み
    products = load_products_json()

    if not products:
        st.info("商品データがありません。")
        return

    # 商品追加モード
    if st.session_state.get("商品追加モード", False):
        st.subheader("➕ 新規商品追加")
        
        with st.form("商品追加フォーム"):
            col1, col2 = st.columns(2)
            
            with col1:
                品名 = st.text_input("品名 *", placeholder="例: 同時通訳者派遣")
                単位 = st.selectbox("単位", ["式", "名", "文字", "ワード", "分", "半日", "日", "時間"])
            
            with col2:
                # min_valueを削除して負の値も入力可能に
                単価 = st.number_input("単価", step=100, help="値引きの場合はマイナス値を入力してください")
                備考 = st.text_input("備考", placeholder="例: 英語、中国語、韓国語も同料金")
            
            # 言語統合のヒント
            st.info("💡 **言語統合のコツ:** 同じサービスで言語が異なる場合は、品名を統一し「備考」欄に対応言語を記載することで商品数を削減できます。明細入力時に具体的な言語を指定してください。")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.form_submit_button("✅ 商品を追加", type="primary"):
                    if not 品名:
                        st.error("品名を入力してください")
                    else:
                        success, message = add_product_to_json(品名, 単位, 単価, 備考)
                        if success:
                            st.success(message)
                            st.session_state["商品追加モード"] = False
                            st.rerun()
                        else:
                            st.error(message)
            
            with col2:
                if st.form_submit_button("❌ キャンセル"):
                    st.session_state["商品追加モード"] = False
                    st.rerun()

    # 商品編集モード
    if st.session_state.get("商品編集モード", False):
        編集中商品 = st.session_state.get("編集中商品")
        if 編集中商品:
            st.subheader(f"📝 商品編集: {編集中商品['品名']}")
            
            with st.form("商品編集フォーム"):
                col1, col2 = st.columns(2)
                
                with col1:
                    品名 = st.text_input("品名 *", value=編集中商品.get("品名", ""))
                    単位選択肢 = ["式", "名", "文字", "ワード", "分", "半日", "日", "時間"]
                    現在の単位 = 編集中商品.get("単位", "式")
                    単位初期index = 単位選択肢.index(現在の単位) if 現在の単位 in 単位選択肢 else 0
                    単位 = st.selectbox("単位", 単位選択肢, index=単位初期index)
                
                with col2:
                    # 単価の安全な処理（負の値も許可）
                    try:
                        現在の単価 = 編集中商品.get("単価", 0)
                        # 数値でない場合は0に設定、負の値は許可
                        if not isinstance(現在の単価, (int, float)):
                            現在の単価 = 0
                        else:
                            現在の単価 = int(現在の単価)
                    except (ValueError, TypeError):
                        現在の単価 = 0
                    
                    # min_valueを削除して負の値も入力可能に
                    単価 = st.number_input("単価", step=100, value=現在の単価, help="値引きの場合はマイナス値を入力してください")
                    備考 = st.text_input("備考", value=編集中商品.get("備考", ""))
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.form_submit_button("✅ 商品を更新", type="primary"):
                        if not 品名:
                            st.error("品名を入力してください")
                        else:
                            success, message = update_product_in_json(編集中商品, 品名, 単位, 単価, 備考)
                            if success:
                                st.success(message)
                                st.session_state["商品編集モード"] = False
                                st.session_state["編集中商品"] = None
                                st.rerun()
                            else:
                                st.error(message)
                
                with col2:
                    if st.form_submit_button("❌ 編集をキャンセル"):
                        st.session_state["商品編集モード"] = False
                        st.session_state["編集中商品"] = None
                        st.rerun()

    # フィルタ適用（内部処理のみ、UI表示なし）
    フィルタ済み商品 = products.copy()

    st.divider()

    # 並び替え操作パネル（単一選択対応版）
    st.subheader("🔄 並び替え・操作")

    # 複数選択・移動・削除機能
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        if st.button("💾 並び順を保存", key="save_order", help="現在の並び順でJSONを保存"):
            if save_products_json(フィルタ済み商品):
                st.success("✅ 並び順を保存しました")
                st.rerun()
            else:
                st.error("❌ 保存に失敗しました")

    with col2:
        # 選択した商品を上に移動（1個以上選択されていれば有効）
        選択商品数 = len(st.session_state.get("selected_products", set()))
        if 選択商品数 >= 1:
            if st.button("⏫ 選択商品を上へ", key="move_selected_up", help=f"選択した{選択商品数}件の商品を上に移動"):
                move_selected_products(フィルタ済み商品, "up")
                st.rerun()
        else:
            st.button("⏫ 選択商品を上へ", disabled=True, help="商品を選択してください")

    with col3:
        # 選択した商品を下に移動（1個以上選択されていれば有効）
        if 選択商品数 >= 1:
            if st.button("⏬ 選択商品を下へ", key="move_selected_down", help=f"選択した{選択商品数}件の商品を下に移動"):
                move_selected_products(フィルタ済み商品, "down")
                st.rerun()
        else:
            st.button("⏬ 選択商品を下へ", disabled=True, help="商品を選択してください")

    with col4:
        # 選択した商品を一括削除（1個以上選択されていれば有効）
        if 選択商品数 >= 1:
            if st.button(f"🗑️ 選択商品を削除 ({選択商品数}件)", key="delete_selected", help=f"選択した{選択商品数}件の商品を一括削除"):
                st.session_state["一括削除確認"] = True
                st.rerun()
        else:
            st.button("🗑️ 選択商品を削除", disabled=True, help="商品を選択してください")

    with col5:
        # 選択解除
        if st.button("🔄 選択解除", key="clear_selection"):
            st.session_state["selected_products"] = set()
            st.rerun()

    # 一括削除確認ダイアログ
    if st.session_state.get("一括削除確認", False):
        選択商品名 = st.session_state.get("selected_products", set())
        st.error(f"⚠️ 選択した{len(選択商品名)}件の商品を削除しますか？")
        st.write("**削除対象:**")
        for 商品名 in sorted(選択商品名):
            st.write(f"- {商品名}")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ はい、一括削除します", key="confirm_batch_delete", type="primary"):
                # 一括削除処理（直接実行）
                try:
                    元リスト = load_products_json()
                    選択商品名 = st.session_state.get("selected_products", set())
                    
                    if 選択商品名:
                        # 選択された商品を除外
                        更新リスト = [p for p in 元リスト if p.get("品名") not in 選択商品名]
                        削除件数 = len(元リスト) - len(更新リスト)
                        
                        if save_products_json(更新リスト):
                            st.success(f"✅ {削除件数}件の商品を削除しました")
                        else:
                            st.error("❌ 一括削除に失敗しました")
                    else:
                        st.warning("削除対象の商品が選択されていません")
                        
                except Exception as e:
                    st.error(f"一括削除処理でエラー: {e}")
                
                st.session_state["一括削除確認"] = False
                st.session_state["selected_products"] = set()
                st.rerun()
        with col2:
            if st.button("❌ キャンセル", key="cancel_batch_delete"):
                st.session_state["一括削除確認"] = False
                st.rerun()

    st.divider()

    if not フィルタ済み商品:
        st.info("商品データがありません。")
        return

    # CSSでコンパクト表示
    st.markdown("""
    <style>
    .compact-table {
        font-size: 14px;
        line-height: 1.2;
    }
    .compact-table .stCheckbox {
        margin: 0;
        padding: 0;
    }
    </style>
    """, unsafe_allow_html=True)

    # 選択状態の初期化
    if "selected_products" not in st.session_state:
        st.session_state["selected_products"] = set()

    # ヘッダー行（固定）
    header_container = st.container()
    with header_container:
        header_col1, header_col2, header_col3, header_col4, header_col5, header_col6, header_col7 = st.columns([0.5, 0.5, 2.5, 0.8, 1.2, 1.5, 1.5])
        header_col1.write("**☑️**")
        header_col2.write("**No**")
        header_col3.write("**品名**")
        header_col4.write("**単位**")
        header_col5.write("**単価**")
        header_col6.write("**備考**")
        header_col7.write("**操作**")

    # 商品一覧をコンパクトに表示
    for i, 商品 in enumerate(フィルタ済み商品):
        with st.container():
            col1, col2, col3, col4, col5, col6, col7 = st.columns([0.5, 0.5, 2.5, 0.8, 1.2, 1.5, 1.5])
            
            # チェックボックス
            with col1:
                商品キー = 商品.get("品名", f"商品{i}")
                checked = st.checkbox("", 
                                    key=f"select_product_{i}",
                                    value=商品キー in st.session_state["selected_products"],
                                    label_visibility="collapsed")
                if checked:
                    st.session_state["selected_products"].add(商品キー)
                else:
                    st.session_state["selected_products"].discard(商品キー)
            
            # 商品情報表示（コンパクト）
            col2.write(f"**{i + 1}**")
            col3.write(f"**{商品.get('品名', '')}**")
            col4.write(商品.get("単位", ""))
            col5.write(f"¥{商品.get('単価', 0):,}")
            
            # 備考の安全な表示（短縮版）
            備考 = 商品.get("備考", "")
            if 備考 and isinstance(備考, str):
                表示備考 = 備考[:15] + "..." if len(備考) > 15 else 備考
            else:
                表示備考 = str(備考)[:15] if 備考 else ""
            col6.write(表示備考)
            
            # 操作ボタン（コンパクト）
            with col7:
                op_col1, op_col2, op_col3, op_col4, op_col5 = st.columns(5)
                
                with op_col1:
                    if st.button("↑", key=f"up_{i}", help="上に移動"):
                        if i > 0:
                            move_single_product(フィルタ済み商品, i, "up")
                            st.rerun()
                
                with op_col2:
                    if st.button("↓", key=f"down_{i}", help="下に移動"):
                        if i < len(フィルタ済み商品) - 1:
                            move_single_product(フィルタ済み商品, i, "down")
                            st.rerun()
                
                with op_col3:
                    if st.button("📝", key=f"edit_{i}", help="編集"):
                        st.session_state["編集中商品"] = 商品
                        st.session_state["商品編集モード"] = True
                        # 編集画面に移動（ページトップにスクロール）
                        st.success(f"商品「{商品['品名']}」の編集画面に移動しました")
                        st.rerun()
                
                with op_col4:
                    if st.button("📋", key=f"add_{i}", help="明細追加"):
                        st.session_state["反映_品名"] = 商品["品名"]
                        st.session_state["反映_単位"] = 商品.get("単位", "式")
                        st.session_state["反映_単価"] = int(商品.get("単価", 0))
                        st.session_state["反映_備考"] = 商品.get("備考", "")
                        st.session_state["アクティブタブ"] = "④ 明細情報を入力"
                        st.success(f"商品「{商品['品名']}」を明細入力に反映しました")
                        st.rerun()
                
                with op_col5:
                    if st.button("🗑️", key=f"del_{i}", help="削除"):
                        st.session_state[f"del_confirm_{i}"] = True
                        st.rerun()
            
            # 削除確認（コンパクト）
            if st.session_state.get(f"del_confirm_{i}", False):
                st.warning(f"「{商品['品名']}」を削除しますか？")
                conf_col1, conf_col2 = st.columns(2)
                with conf_col1:
                    if st.button("削除", key=f"conf_del_{i}"):
                        # 商品を削除
                        元リスト = load_products_json()
                        元リスト = [p for p in 元リスト if p.get("品名") != 商品["品名"]]
                        if save_products_json(元リスト):
                            st.success(f"商品「{商品['品名']}」を削除しました")
                            st.session_state[f"del_confirm_{i}"] = False
                            st.rerun()
                        else:
                            st.error("削除に失敗しました")
                with conf_col2:
                    if st.button("キャンセル", key=f"cancel_del_{i}"):
                        st.session_state[f"del_confirm_{i}"] = False
                        st.rerun()

def load_customers_json():
    """顧客JSONファイルを読み込む"""
    try:
        customers_json_file = os.path.join(DATA_FOLDER, "customers.json")
        if os.path.exists(customers_json_file):
            with open(customers_json_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        else:
            return []
    except Exception as e:
        st.error(f"顧客データの読み込みエラー: {e}")
        return []

def save_customers_json(customers_list):
    """顧客データをJSONファイルに保存"""
    try:
        os.makedirs(DATA_FOLDER, exist_ok=True)
        customers_json_file = os.path.join(DATA_FOLDER, "customers.json")
        with open(customers_json_file, "w", encoding="utf-8") as f:
            json.dump(customers_list, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"顧客データの保存エラー: {e}")
        return False

def add_customer_to_json(顧客会社名, 顧客部署名, 顧客担当者, 郵便番号, 住所1, 住所2):
    """新規顧客をJSONに追加（住所細分化対応・修正版）"""
    customers = load_customers_json()

    # 重複チェック（会社名、部署名、担当者の組み合わせ）
    for customer in customers:
        if (customer.get("顧客会社名") == 顧客会社名 and 
            customer.get("顧客部署名") == 顧客部署名 and 
            customer.get("顧客担当者") == 顧客担当者):
            return False, "同じ顧客情報が既に登録されています"

    # 顧客Noの生成（会社名ベース）
    顧客No = 1
    if customers:
        # 同じ会社名の顧客を検索
        同一会社の顧客 = [c for c in customers if c.get("顧客会社名") == 顧客会社名]
        if 同一会社の顧客:
            # 同一会社の場合は既存の顧客Noを使用
            顧客No = 同一会社の顧客[0].get("顧客No", 1)
        else:
            # 新しい会社の場合は最大顧客No + 1
            existing_companies = list(set([c.get("顧客会社名") for c in customers]))
            顧客No = len(existing_companies) + 1

    # 統合住所の生成（互換性のため）
    統合住所 = f"{郵便番号} {住所1} {住所2}".strip()

    # 新規顧客データ（住所細分化対応・統合住所追加）
    new_customer = {
        "顧客No": 顧客No,
        "顧客会社名": 顧客会社名,
        "顧客部署名": 顧客部署名,
        "顧客担当者": 顧客担当者,
        "郵便番号": 郵便番号,
        "住所1": 住所1,
        "住所2": 住所2,
        "顧客住所": 統合住所,  # 統合住所も保存（互換性のため）
        "登録日": datetime.date.today().strftime("%Y-%m-%d")
    }

    customers.append(new_customer)
    
    # 顧客Noと会社名で自動並び替え
    customers.sort(key=lambda x: (x.get("顧客No", 999), x.get("顧客会社名", ""), x.get("顧客担当者", "")))

    if save_customers_json(customers):
        return True, f"顧客「{顧客会社名}」を登録しました"
    else:
        return False, "顧客データの保存に失敗しました"

def render_customer_list_tab():
    """顧客一覧タブを表示（住所表示修正版）"""
    st.header("⑤ 顧客一覧")

    # 顧客データの読み込み
    customers = load_customers_json()

    if not customers:
        st.info("顧客データがありません。")
        
        st.divider()
        if st.button("新しい顧客を手動追加"):
            st.session_state["アクティブタブ"] = "② 顧客情報を入力"
            st.rerun()
        return

    # 検索・フィルタ機能
    st.subheader("🔍 検索・フィルタ")

    col1, col2 = st.columns(2)

    with col1:
        検索キーワード = st.text_input("🔍 顧客名で検索", placeholder="会社名・部署名・担当者名を入力", key="customer_search")

    with col2:
        # 会社でフィルタ
        会社リスト = ["すべて"] + sorted(list(set([c["顧客会社名"] for c in customers if c["顧客会社名"]])))
        選択された会社 = st.selectbox("会社で絞り込み", 会社リスト, key="customer_company_filter")

    # フィルタ適用
    フィルタ済み顧客 = customers.copy()

    if 選択された会社 != "すべて":
        フィルタ済み顧客 = [c for c in フィルタ済み顧客 if c["顧客会社名"] == 選択された会社]

    if 検索キーワード:
        フィルタ済み顧客 = [
            c for c in フィルタ済み顧客 
            if (検索キーワード in c.get("顧客会社名", "") or 
                検索キーワード in c.get("顧客部署名", "") or 
                検索キーワード in c.get("顧客担当者", ""))
        ]

    # 顧客Noでソート（会社名でグループ化するため）
    フィルタ済み顧客.sort(key=lambda x: (x.get("顧客No", 999), x.get("顧客会社名", ""), x.get("顧客担当者", "")))

    # 会社名でグループ化
    会社別顧客 = {}
    for 顧客 in フィルタ済み顧客:
        会社名 = 顧客["顧客会社名"]
        if 会社名 not in 会社別顧客:
            会社別顧客[会社名] = []
        会社別顧客[会社名].append(顧客)

    # 統計情報
    st.subheader("📊 統計情報")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("総顧客数", len(customers))

    with col2:
        会社数 = len(set([c["顧客会社名"] for c in customers]))
        st.metric("登録会社数", 会社数)

    with col3:
        st.metric("表示件数", len(フィルタ済み顧客))

    st.divider()

    # 顧客一覧の表示（住所表示修正版）
    st.subheader(f"👥 顧客一覧 ({len(フィルタ済み顧客)}件)")

    if not フィルタ済み顧客:
        st.info("条件に合致する顧客がありません。")
        return

    # 会社別に表示
    全体カウンタ = 0  # 一意のキー生成用
    
    for 会社名, 会社の顧客リスト in 会社別顧客.items():
        # 会社名を表示
        代表顧客 = 会社の顧客リスト[0]
        顧客No = 代表顧客.get("顧客No", "未設定")
        担当者数 = len(会社の顧客リスト)
        
        # 会社情報をexpanderで表示
        with st.expander(f"🏢 [{顧客No}] {会社名} ({担当者数}名)", expanded=False):
            
            # 会社の住所情報を表示（新形式優先、旧形式フォールバック）
            郵便番号 = 代表顧客.get("郵便番号", "")
            住所1 = 代表顧客.get("住所1", "")
            住所2 = 代表顧客.get("住所2", "")
            旧住所 = 代表顧客.get("顧客住所", "")
            
            # 新形式の住所があれば表示
            if 郵便番号 or 住所1 or 住所2:
                住所表示 = f"{郵便番号} {住所1} {住所2}".strip()
                st.info(f"📍 住所: {住所表示}")
            elif 旧住所:
                # 旧形式の住所のみある場合
                st.info(f"📍 住所: {旧住所} （旧形式）")
                st.warning("💡 編集して新形式に移行することをお勧めします")
            else:
                st.warning("📍 住所が未設定です")
            
            # 担当者一覧をテーブル形式で表示
            st.write("**👤 担当者一覧**")
            
            for i, 顧客 in enumerate(会社の顧客リスト):
                全体カウンタ += 1
                
                # 担当者情報をコンテナで区切って表示
                with st.container():
                    st.write(f"**{i+1}. {顧客['顧客担当者']}** ({顧客.get('顧客部署名', '部署未設定')})")
                    
                    # 詳細情報を折りたたみ形式で表示（セッション状態を使用）
                    詳細表示キー = f"顧客詳細_{全体カウンタ}"
                    詳細表示状態 = st.session_state.get(詳細表示キー, False)
                    
                    col1, col2 = st.columns([1, 6])
                    with col1:
                        if st.button("📋 詳細", key=f"show_detail_{全体カウンタ}"):
                            st.session_state[詳細表示キー] = not 詳細表示状態
                            st.rerun()
                    
                    with col2:
                        # 操作ボタンを横並びで配置
                        btn_col1, btn_col2, btn_col3 = st.columns(3)
                        
                        with btn_col1:
                            if st.button("📝 編集", key=f"edit_customer_{全体カウンタ}"):
                                st.session_state["編集中顧客"] = 顧客
                                st.session_state["アクティブタブ"] = "② 顧客情報を入力"
                                st.rerun()
                        
                        with btn_col2:
                            if st.button("📋 案件作成", key=f"create_project_{全体カウンタ}"):
                                # 顧客情報をセッションに設定（住所移行対応）
                                郵便番号 = 顧客.get("郵便番号", "")
                                住所1 = 顧客.get("住所1", "")
                                住所2 = 顧客.get("住所2", "")
                                
                                # 新形式の住所がない場合は旧住所を分割
                                if not (郵便番号 or 住所1 or 住所2):
                                    旧住所 = 顧客.get("顧客住所", "")
                                    if 旧住所:
                                        from estimate_excel_writer import parse_address
                                        郵便番号, 住所1, 住所2 = parse_address(旧住所)
                                
                                st.session_state["選択された顧客会社名"] = 顧客["顧客会社名"]
                                st.session_state["選択された顧客部署名"] = 顧客.get("顧客部署名", "")
                                st.session_state["選択された顧客担当者"] = 顧客["顧客担当者"]
                                st.session_state["選択された郵便番号"] = 郵便番号
                                st.session_state["選択された住所1"] = 住所1
                                st.session_state["選択された住所2"] = 住所2
                                st.session_state["アクティブタブ"] = "③ 案件情報を入力"
                                st.success(f"顧客「{顧客['顧客会社名']} - {顧客['顧客担当者']}」で案件作成を開始します")
                                st.rerun()
                        
                        with btn_col3:
                            if st.button("🗑️ 削除", key=f"delete_customer_{全体カウンタ}"):
                                st.session_state[f"削除確認_顧客_{全体カウンタ}"] = True
                                st.rerun()
                    
                    # 詳細情報の表示（セッション状態で制御、住所表示修正）
                    if st.session_state.get(詳細表示キー, False):
                        detail_col1, detail_col2 = st.columns(2)
                        
                        with detail_col1:
                            st.write(f"**顧客No:** {顧客.get('顧客No', '未設定')}")
                            st.write(f"**会社名:** {顧客['顧客会社名']}")
                            st.write(f"**部署名:** {顧客.get('顧客部署名', '未設定')}")
                        
                        with detail_col2:
                            # 住所の詳細表示（新形式優先）
                            個別郵便番号 = 顧客.get("郵便番号", "")
                            個別住所1 = 顧客.get("住所1", "")
                            個別住所2 = 顧客.get("住所2", "")
                            個別旧住所 = 顧客.get("顧客住所", "")
                            
                            if 個別郵便番号 or 個別住所1 or 個別住所2:
                                個別住所表示 = f"{個別郵便番号} {個別住所1} {個別住所2}".strip()
                                st.write(f"**住所:** {個別住所表示}")
                            elif 個別旧住所:
                                st.write(f"**住所:** {個別旧住所} （旧形式）")
                            else:
                                st.write("**住所:** 未設定")
                            
                            st.write(f"**登録日:** {顧客.get('登録日', '未設定')}")
                            if 顧客.get("更新日"):
                                st.write(f"**更新日:** {顧客['更新日']}")
                    
                    # 削除確認ダイアログ
                    if st.session_state.get(f"削除確認_顧客_{全体カウンタ}", False):
                        st.warning(f"顧客「{顧客['顧客会社名']} - {顧客['顧客担当者']}」を本当に削除しますか？")
                        confirm_col1, confirm_col2 = st.columns(2)
                        
                        with confirm_col1:
                            if st.button("はい、削除します", key=f"confirm_delete_customer_{全体カウンタ}"):
                                # 顧客を削除
                                customers_updated = load_customers_json()  # 最新データを再読み込み
                                削除対象 = None
                                for customer in customers_updated:
                                    if (customer.get("顧客会社名") == 顧客["顧客会社名"] and 
                                        customer.get("顧客部署名") == 顧客.get("顧客部署名", "") and 
                                        customer.get("顧客担当者") == 顧客["顧客担当者"]):
                                        削除対象 = customer
                                        break
                                
                                if 削除対象:
                                    customers_updated.remove(削除対象)
                                    if save_customers_json(customers_updated):
                                        st.success(f"顧客「{顧客['顧客会社名']} - {顧客['顧客担当者']}」を削除しました")
                                        # セッション状態のクリーンアップ
                                        keys_to_clean = [key for key in st.session_state.keys() 
                                                    if key.startswith(f"削除確認_顧客_")]
                                        for key in keys_to_clean:
                                            if key in st.session_state:
                                                del st.session_state[key]
                                        st.rerun()
                                    else:
                                        st.error("削除に失敗しました")
                                else:
                                    st.error("削除対象の顧客が見つかりませんでした")
                        
                        with confirm_col2:
                            if st.button("キャンセル", key=f"cancel_delete_customer_{全体カウンタ}"):
                                st.session_state[f"削除確認_顧客_{全体カウンタ}"] = False
                                st.rerun()
                    
                    # 区切り線
                    if i < len(会社の顧客リスト) - 1:
                        st.write("---")

    # 新規顧客追加ボタン
    st.divider()
    if st.button("➕ 新しい顧客を追加", type="primary"):
        st.session_state["アクティブタブ"] = "② 顧客情報を入力"
        st.rerun()

# 商品データのJSON管理関数
def load_products_json():
    """商品JSONファイルを読み込む"""
    try:
        products_json_file = os.path.join(DATA_FOLDER, "products.json")
        if os.path.exists(products_json_file):
            with open(products_json_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        else:
            return []
    except Exception as e:
        st.error(f"商品データの読み込みエラー: {e}")
        return []

def save_products_json(products_list):
    """商品データをJSONファイルに保存"""
    try:
        os.makedirs(DATA_FOLDER, exist_ok=True)
        products_json_file = os.path.join(DATA_FOLDER, "products.json")
        with open(products_json_file, "w", encoding="utf-8") as f:
            json.dump(products_list, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"商品データの保存エラー: {e}")
        return False

def add_product_to_json(品名, 単位, 単価, 備考):
    """新規商品をJSONに追加（移行元フィールド削除版）"""
    products = load_products_json()

    # 重複チェック
    for product in products:
        if product.get("品名") == 品名:
            return False, "同じ商品名が既に登録されています"

    # 新規商品データ（移行元フィールドを削除）
    new_product = {
        "品名": 品名,
        "単位": 単位,
        "単価": float(単価),
        "備考": 備考,
        "登録日": datetime.date.today().strftime("%Y-%m-%d")
    }

    products.append(new_product)

    if save_products_json(products):
        return True, f"商品「{品名}」を登録しました"
    else:
        return False, "商品データの保存に失敗しました"

def update_product_in_json(元商品データ, 品名, 単位, 単価, 備考):
    """商品情報をJSONで更新"""
    try:
        products = load_products_json()
        
        # 元の商品データを検索
        for i, product in enumerate(products):
            if product.get("品名") == 元商品データ["品名"]:
                
                # 重複チェック（自分以外で同じ商品名）
                if 品名 != 元商品データ["品名"]:  # 商品名が変更された場合のみチェック
                    for j, other_product in enumerate(products):
                        if i != j and other_product.get("品名") == 品名:
                            return False, "同じ商品名が既に存在します"
                
                # 商品情報を更新
                products[i]["品名"] = 品名
                products[i]["単位"] = 単位
                products[i]["単価"] = float(単価)
                products[i]["備考"] = 備考
                products[i]["更新日"] = datetime.date.today().strftime("%Y-%m-%d")
                
                if save_products_json(products):
                    return True, f"商品「{品名}」を更新しました"
                else:
                    return False, "商品データの保存に失敗しました"
        
        return False, "更新対象の商品が見つかりませんでした"
        
    except Exception as e:
        return False, f"更新処理でエラーが発生しました: {e}"

# サイドバーのボタン修正版（新規案件作成ボタンをサイドバーに移動）
def render_sidebar_status():
    """サイドバーに現在の状態を表示（分類項目除外版）"""
    with st.sidebar:
        st.header("現在の状態")
    
        # 顧客情報
        st.subheader("顧客情報")
        顧客会社名 = st.session_state.get("選択された顧客会社名", "未選択")
        顧客部署名 = st.session_state.get("選択された顧客部署名", "未選択")
        顧客担当者 = st.session_state.get("選択された顧客担当者", "未選択")
        
        st.write(f"**会社名:** {顧客会社名}")
        st.write(f"**部署名:** {顧客部署名}")
        st.write(f"**担当者:** {顧客担当者}")
        
        # 案件情報
        st.subheader("案件情報")
        案件名 = st.session_state.get("案件名", "未入力")
        見積No = st.session_state.get("見積No", "未生成")
        発行日 = st.session_state.get("発行日", "未設定")
        
        st.write(f"**案件名:** {案件名}")
        st.write(f"**見積No:** {見積No}")
        st.write(f"**発行日:** {発行日}")
        
        # 明細情報（分類項目を除外して計算）
        st.subheader("明細情報")
        全明細リスト = st.session_state.get("明細リスト", [])
        通常明細リスト = [item for item in 全明細リスト if not item.get("分類", False)]
        
        明細数 = len(通常明細リスト)
        合計金額 = 0
        for item in 通常明細リスト:
            金額 = item.get("金額", 0)
            if isinstance(金額, (int, float)):
                合計金額 += 金額
        
        st.write(f"**明細数:** {明細数}件")
        st.write(f"**合計金額:** ¥{合計金額:,}")

        # 操作ボタン
        st.subheader("操作")
        if st.button("データをリセット", key="reset_button"):
            st.session_state["リセット確認中"] = True

        # 新規案件作成ボタンをここに移動
        if st.button("➕ 新しい案件を作成", type="primary", key="create_new_project_sidebar"):
            # データを自動リセット
            reset_all_data()
            st.session_state["アクティブタブ"] = "② 顧客情報を入力"
            st.success("新しい案件を作成します。データをリセットしました。")
            st.rerun()

        # 確認UIの表示制御
        if st.session_state.get("リセット確認中", False):
            st.warning("本当にリセットしますか？")
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("はい", key="confirm_reset"):
                    reset_all_data()
                    st.session_state["リセット確認中"] = False
                    st.rerun()
            with col2:
                if st.button("キャンセル", key="cancel_reset"):
                    st.session_state["リセット確認中"] = False
        
        if 明細数 > 0 and st.button("見積書を出力"):
            export_estimate()

def reset_all_data():
    """すべてのデータをリセット（発行日追加版）"""
    keys_to_reset = [
        "明細リスト", "編集対象", "案件名", "見積No", "発行日",  # ← 発行日を追加
        "選択された顧客会社名", "選択された顧客部署名", "選択された顧客担当者", "選択された顧客住所",
        "選択された郵便番号", "選択された住所1", "選択された住所2",
        "発行者名", "備考", "メモ", "状況", "受注日", "納品日",
        "売上額", "仕入額", "粗利", "粗利率", "売上額自動更新", "担当部署"
    ]

    # 入力中データもクリア（住所関連も追加）
    input_keys_to_clear = [
        "入力中_顧客会社名選択", "入力中_顧客会社名", "入力中_顧客担当者選択",
        "入力中_顧客担当者", "入力中_顧客部署名", "入力中_顧客住所",
        "入力中_郵便番号", "入力中_住所1", "入力中_住所2",  # 住所細分化入力中データも追加
        "入力中_品名", "入力中_数量", "入力中_単位", "入力中_単価", "入力中_備考",
        "前回選択品名", "リセット確認中", "編集中顧客", "住所補完済み"  # 住所補完フラグも追加
    ]

    # 削除確認フラグもクリア
    削除確認Keys = [key for key in st.session_state.keys() if key.startswith("削除確認_")]

    for key in keys_to_reset:
        if key == "明細リスト":
            st.session_state[key] = []
        elif key == "発行日":
            st.session_state[key] = None  # ← ここを修正（datetime.date.today() から None に変更）
        elif key == "発行者名":
            st.session_state[key] = ISSUER_LIST[0]
        elif key == "売上額自動更新":
            st.session_state[key] = True
        elif key == "状況":
            st.session_state[key] = "見積中"
        elif key in ["受注日", "納品日"]:
            st.session_state[key] = None
        elif key in ["売上額", "仕入額", "粗利"]:
            st.session_state[key] = 0
        elif key == "粗利率":
            st.session_state[key] = 0
        else:
            st.session_state[key] = ""

    # 入力中データを削除
    for key in input_keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]

    # 削除確認フラグを削除
    for key in 削除確認Keys:
        if key in st.session_state:
            del st.session_state[key]

    # アクティブタブを確実に設定
    st.session_state["アクティブタブ"] = "① 案件一覧"
    st.success("データをリセットしました")

def apply_custom_css():
    """カスタムCSSを適用して表示を改善"""
    st.markdown("""
    <style>
    /* expanderヘッダーの文字サイズを拡大 */
    .streamlit-expanderHeader {
        font-size: 16px !important;
        font-weight: bold !important;
        line-height: 1.4 !important;
    }
    
    /* expanderヘッダー内のテキストサイズ調整 */
    .streamlit-expanderHeader p {
        font-size: 16px !important;
        margin: 0 !important;
    }
    
    /* expanderアイコンのサイズ調整 */
    .streamlit-expanderHeader svg {
        width: 20px !important;
        height: 20px !important;
    }
    
    /* 一般的なテキストサイズを少し大きく */
    .stMarkdown p {
        font-size: 15px !important;
    }
    
    /* メトリクス表示の改善 */
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

# ログイン認証関数
def authenticate_user(username, password):
    """ユーザー認証を行う"""
    # 認証情報（実際の運用では環境変数やデータベースを使用）
    USERS = {
        "admin": "password123",
        "user1": "pass1234",
        "manager": "mgr2024",
        "staff": "staff123"
    }
    
    return USERS.get(username) == password

def render_login_form():
    """ログインフォームを表示"""
    st.title("🔐 見積書作成アプリ - ログイン")
    
    # ログインフォームを中央に配置
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### ログインしてください")
        
        with st.form("login_form"):
            username = st.text_input("ユーザー名", placeholder="ユーザー名を入力")
            password = st.text_input("パスワード", type="password", placeholder="パスワードを入力")
            
            login_button = st.form_submit_button("ログイン", type="primary", use_container_width=True)
            
            if login_button:
                if not username or not password:
                    st.error("ユーザー名とパスワードを入力してください")
                elif authenticate_user(username, password):
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username
                    st.session_state["login_time"] = datetime.datetime.now()
                    st.success("ログインしました！")
                    st.rerun()
                else:
                    st.error("ユーザー名またはパスワードが正しくありません")
        
        # デモ用認証情報の表示
        st.markdown("---")
        st.info("**デモ用ログイン情報:**\n\n"
               "• ユーザー名: `admin` / パスワード: `password123`\n\n"
               "• ユーザー名: `user1` / パスワード: `pass1234`\n\n"
               "• ユーザー名: `manager` / パスワード: `mgr2024`\n\n"
               "• ユーザー名: `staff` / パスワード: `staff123`")

def render_logout_button():
    """ログアウトボタンをサイドバーに表示"""
    with st.sidebar:
        st.markdown("---")
        st.write(f"**ログイン中:** {st.session_state.get('username', 'Unknown')}")
        
        # ログイン時間の表示
        login_time = st.session_state.get('login_time')
        if login_time:
            elapsed = datetime.datetime.now() - login_time
            hours, remainder = divmod(elapsed.total_seconds(), 3600)
            minutes, _ = divmod(remainder, 60)
            st.write(f"**ログイン時間:** {int(hours)}時間{int(minutes)}分")
        
        if st.button("🚪 ログアウト", type="secondary", use_container_width=True):
            # セッション状態をクリア
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.success("ログアウトしました")
            st.rerun()

def check_session_timeout():
    """セッションタイムアウトをチェック（オプション）"""
    if "login_time" in st.session_state:
        login_time = st.session_state["login_time"]
        elapsed = datetime.datetime.now() - login_time
        
        # 8時間でタイムアウト
        if elapsed.total_seconds() > 8 * 3600:
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.warning("セッションがタイムアウトしました。再度ログインしてください。")
            st.rerun()

def main():
    """メイン処理（ログイン機能付き）"""
    # セッションタイムアウトチェック
    check_session_timeout()
    
    # 認証チェック
    if not st.session_state.get("authenticated", False):
        render_login_form()
        return
    
    # ログイン後のメイン処理
    # カスタムCSSの適用
    apply_custom_css()
    
    # セッション状態の初期化
    init_session_state()

    # データの読み込み（JSONから）
    顧客一覧, _, 品名一覧 = load_data()
        
    # タブの選択
    タブ選択肢 = ["① 案件一覧", "② 顧客情報を入力", "③ 案件情報を入力", "④ 明細情報を入力", "⑤ 顧客一覧", "⑥ 商品一覧"]

    # アクティブタブが選択肢にない場合はデフォルトに設定
    現在のタブ = st.session_state.get("アクティブタブ", "① 案件一覧")
    if 現在のタブ not in タブ選択肢:
        現在のタブ = "① 案件一覧"
        st.session_state["アクティブタブ"] = 現在のタブ

    # タブ切り替えをセッション状態と同期
    タブ = st.radio(
        "手順を選択してください",
        タブ選択肢,
        index=タブ選択肢.index(現在のタブ),
        key="main_tab_radio"
    )

    # タブが変更された場合のみセッション状態を更新
    if タブ != st.session_state.get("アクティブタブ"):
        st.session_state["アクティブタブ"] = タブ

    # 各タブの処理
    if タブ == "① 案件一覧":
        render_project_list_tab()
    elif タブ == "② 顧客情報を入力":
        render_customer_tab(顧客一覧)
    elif タブ == "③ 案件情報を入力":
        render_project_tab()
    elif タブ == "④ 明細情報を入力":
        render_detail_tab(品名一覧)
    elif タブ == "⑤ 顧客一覧":
        render_customer_list_tab()
    elif タブ == "⑥ 商品一覧":
        render_product_list_tab()

    # サイドバーに現在の状態を表示（ログアウトボタン付き）
    render_sidebar_status()
    render_logout_button()

# アプリケーションの実行
if __name__ == "__main__":
    main()