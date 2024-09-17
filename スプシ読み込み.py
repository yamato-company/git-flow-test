import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import argparse

# Google Sheets API の設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = 'C:/Users/81904/Downloads/test-435704-7da9856d593e.json'
SPREADSHEET_ID = '1Ovjy1nopeScqyytL3yEV1EPGn9X2BWYnKyuV8V3N4-g'
SHEET_NAME = 'エンド管理'

def main(target_date=None):
    try:
        # サービスアカウントの認証情報を取得
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        # Sheets API クライアントを構築
        service = build('sheets', 'v4', credentials=creds)

        # スプレッドシートからデータを取得
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=f'{SHEET_NAME}!A:O').execute()
        values = result.get('values', [])

        if not values:
            print('データが見つかりません。')
            return

        # データフレームに変換
        df = pd.DataFrame(values[1:], columns=values[0])

        # 列名を確認
        print("データフレームの列名:")
        print(df.columns)

        # 必要な列が存在するか確認
        required_columns = ['企業名(スペース削除で統一)', '施設名(スペース削除で統一)', '職種', '更新日', 'ステータス']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"警告: 以下の列が見つかりません: {', '.join(missing_columns)}")
            return

        # 指定された日付または今日の日付を使用
        if target_date:
            date_to_use = target_date
        else:
            date_to_use = datetime.now().strftime('%Y/%m/%d')
        
        print(f"使用する日付: {date_to_use}")

        # 条件に合うデータを抽出
        filtered_df = df[
            (df['更新日'] == date_to_use) & 
            (df['ステータス'].isin(['募集中', '急募', '都度紹介OK']))
        ][['企業名(スペース削除で統一)', '施設名(スペース削除で統一)', '職種']]

        # 列名を簡略化
        filtered_df.columns = ['企業名', '施設名', '職種']

        # 結果を表示
        if filtered_df.empty:
            print('条件に合うデータが見つかりません。')
        else:
            print("抽出されたデータ:")
            print(filtered_df)

    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='スプレッドシートからデータを抽出します。')
    parser.add_argument('--date', type=str, help='処理する日付（YYYY/MM/DD形式）。指定しない場合は今日の日付が使用されます。')
    args = parser.parse_args()

    main(args.date)