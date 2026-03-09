#!/usr/bin/env python3
"""
AWS Summit Japan 2026 パートナーセッション
mdの仕様から直接JSONを構築しPPTXを生成するスクリプト
"""
import json
from generate_ppt import create_pptx

presentation_data = {
    "title": "セキュアクラウド環境での音声基盤を提供するNotta ― 次なる業務効率化の進化",
    "subtitle": "AWS Summit Japan 2026 | パートナーセッション",
    "slides": [
        # ── スライド1: 表紙 ──
        {
            "type": "title",
            "title": "セキュアクラウド環境での音声基盤を提供するNotta\n― 次なる業務効率化の進化",
            "content": "AWS Summit Japan 2026 パートナーセッション\n河野 秀彰（FTC）/ 華 逸東（FDE）\n2026年6月25日〜26日 | 幕張メッセ"
        },
        # ── スライド2: アジェンダ ──
        {
            "type": "content",
            "slide_title": "本日のアジェンダ",
            "key_message": "セキュアな音声基盤構築から、LLM×音声による業務自動化の実践まで",
            "body_items": [
                {"heading": "1. Notta サービス紹介", "body": "AI文字起こし・要約プラットフォームの概要と導入実績"},
                {"heading": "2. セキュリティ要件", "body": "エンタープライズが求める音声基盤のセキュリティ要件を解説"},
                {"heading": "3. セキュアアーキテクチャ", "body": "AWS VPC + PrivateLink を活用したセキュア音声基盤の設計"},
                {"heading": "4. LLM × 音声の業務自動化", "body": "Bedrock Agents × Notta で実現するライブデモと今後の展望"}
            ],
            "notes": "40分のセッション。前半15分: 河野（セキュア基盤）、後半15分: 華（LLM業務自動化デモ）"
        },
        # ── スライド3: Notta サービス紹介 ──
        {
            "type": "content",
            "slide_title": "Notta ― AI文字起こし・要約プラットフォーム",
            "key_message": "104言語対応のリアルタイム文字起こしと、SOC 2 / ISO 27001 / HIPAA 認証取得のセキュリティ基盤",
            "body_items": [
                {"heading": "リアルタイム文字起こし", "body": "104言語対応。Web会議Bot自動入室（Zoom / Teams / Google Meet / Webex）で手間なく記録"},
                {"heading": "AI要約・翻訳自動生成", "body": "会議終了後に自動でサマリー生成。多言語翻訳にも対応し、グローバルチームの情報共有を効率化"},
                {"heading": "マルチプラットフォーム", "body": "Web / モバイルアプリ / Chrome拡張 / API。あらゆる利用シーンに対応した柔軟な提供形態"},
                {"heading": "セキュリティ認証", "body": "SOC 2 Type II / ISO 27001 / HIPAA 取得済み。エンタープライズ水準のコンプライアンスに対応"}
            ],
            "notes": "Nottaの基本機能を3分で説明。利用シーンのフロー図（Web会議→文字起こし→要約→共有）を見せながら"
        },
        # ── スライド4: エンタープライズの音声データ課題 ──
        {
            "type": "content",
            "slide_title": "なぜ音声基盤にセキュリティが求められるのか",
            "key_message": "音声データ活用の加速とセキュリティ要求の高まり ― 「便利さ」と「安全性」の両立が課題",
            "body_items": [
                {"heading": "音声データ活用の加速", "body": "ハイブリッドワーク定着でWeb会議が増加。議事録・商談記録のデジタル化とAI活用による意思決定高速化が進む"},
                {"heading": "金融・官公庁の要求", "body": "金融機関: FISC安全対策基準\n官公庁: 政府統一基準群（ISMAP）\nグローバル基準への準拠が必須"},
                {"heading": "医療・一般企業の要求", "body": "医療機関: HIPAA準拠\n一般企業: 社内情報の外部流出リスク対策\nデータ主権の確保が重要テーマに"}
            ],
            "notes": "業界別のセキュリティ要件を比較表で示す。2分で説明"
        },
        # ── スライド5: インフラストラクチャ構成 ──
        {
            "type": "content",
            "slide_title": "AWS上に構築されたNottaのセキュアインフラ",
            "key_message": "東京リージョン・N+1冗長構成・SLA 99% ― AES128二次暗号化とKMS鍵管理による多層防御",
            "body_items": [
                {"heading": "データ保管", "body": "音声・録画: Amazon S3（SSE暗号化）\n構造化データ: Amazon Aurora\n暗号鍵管理: AWS KMS"},
                {"heading": "通信の暗号化", "body": "HTTPS/TLS トランスポート暗号化\nAES128 二次暗号化\nパスワード: scram-sha-256 ハッシュ"},
                {"heading": "冗長構成・バックアップ", "body": "コンピューティング: N+1冗長構成\nバックアップ: RPO 1日前 / 半年間保存 / 35世代\nSLA 99%"},
                {"heading": "東京リージョン", "body": "AWS 東京リージョン（国内データセンター）でのデータ保管。日本のデータ主権要件に対応"}
            ],
            "notes": "AWSアイコンを使ったインフラ構成図を表示しながら説明。3分"
        },
        # ── スライド6: データフローとAI学習制御 ──
        {
            "type": "content",
            "slide_title": "メモリバッファ処理 ― エンジン側にデータは残らない",
            "key_message": "音声認識はメモリバッファ処理で即削除。AWS Bedrock経由のAI要約もデータ学習・保管なし",
            "body_items": [
                {"heading": "Step 1-2: 音声入力", "body": "利用者が録音 or Web会議URLを入力\n音声データをSSL/TLS経由で音声認識エンジンに送信"},
                {"heading": "Step 3-4: 文字起こし", "body": "音声認識エンジンにてメモリバッファ処理→即削除\n文字起こしテキストをお客様S3領域に保存"},
                {"heading": "Step 5-6: AI要約", "body": "AWS Bedrock（Claude）経由で要約処理\nデータ学習・保管は一切なし\n生成された要約ファイルをS3に保存"},
                {"heading": "AI学習制御", "body": "「AI学習なし」（エンタープライズ推奨）: データ完全非保持\n「AI学習あり」: 一部ランダム学習（選択可能）"}
            ],
            "notes": "データフロー図を使ってステップバイステップで説明。3分。メモリバッファ処理の安全性を強調"
        },
        # ── スライド7: VPC接続 ──
        {
            "type": "content",
            "slide_title": "Amazon VPC + PrivateLink によるセキュアな音声基盤アーキテクチャ",
            "key_message": "パブリックインターネットを経由しない閉域網設計 ― VPCエンドポイント + 多層防御",
            "body_items": [
                {"heading": "閉域網接続の設計思想", "body": "パブリックインターネットを経由しないデータ通信\nAWS PrivateLink を活用した VPC エンドポイント構成"},
                {"heading": "VPC構成要素", "body": "プライベートサブネット構成\nVPCエンドポイント: bedrock-runtime\nDirect Connect / VPN 経由のオンプレ接続"},
                {"heading": "多層防御", "body": "セキュリティグループ + NACLs\nIAMエンドポイントポリシーによるアクセス制御\nプライベートDNS有効化"},
                {"heading": "Bedrock AgentCore VPC対応", "body": "Runtime / Browser / Code Interpreter のVPC接続\n東京リージョン対応済み\nエンタープライズ向け閉域実行環境"}
            ],
            "notes": "VPC構成のアーキテクチャ図を表示。左: オンプレ→DC/VPN→VPC、中央: プライベートサブネット、右: Bedrock/S3。3分"
        },
        # ── スライド8: エンタープライズ向けセキュリティ機能 ──
        {
            "type": "content",
            "slide_title": "多層的なセキュリティ ― SSO / IP制限 / 監査ログ",
            "key_message": "既存IT基盤との統合と、きめ細かなアクセス制御で、エンタープライズのセキュリティ要件に対応",
            "body_items": [
                {"heading": "SSO連携", "body": "SAML SSO 対応（Entra ID / Okta / OneLogin）\n既存IdPとの統合によるアカウント管理一元化"},
                {"heading": "IPアドレス制限", "body": "指定グローバルIPからのアクセスのみ許可\nVPN接続でリモートワーク対応\nエンタープライズ: 最大1,000個のIP登録"},
                {"heading": "管理・監査機能", "body": "操作ログ・監査ログ（ログイン / ノート操作 / DL履歴）\n外部共有制御 / 権限管理 / 利用状況レポート"}
            ],
            "notes": "エンタープライズ導入の全体イメージ図: お客様IT基盤→Notta→AWS日本国内DC。2分"
        },
        # ── スライド9: 外部サービス連携 ──
        {
            "type": "content",
            "slide_title": "入力から出力まで ― シームレスなデータ活用",
            "key_message": "会議プラットフォームからCRM・ドキュメントまで、収録→要約→外部送信を完全自動化",
            "body_items": [
                {"heading": "入力連携", "body": "Zoom / Teams / Google Meet / Webex\nNotta Botが自動入室し文字起こし・録画"},
                {"heading": "スケジュール連携", "body": "Google カレンダー / Outlook カレンダー\n予定会議にBotを自動派遣"},
                {"heading": "出力先連携", "body": "CRM: Salesforce / HubSpot\nドキュメント: Google Docs / Notion\nストレージ: Google Drive\n通知: Slack"},
                {"heading": "自動化ルール", "body": "収録完了→AI要約自動生成→外部サービスへ自動送信\nZapier経由でさらに拡張可能"}
            ],
            "notes": "入力→Notta処理→出力の3レイヤーフロー図を表示。2分"
        },
        # ── スライド10: 後半導入 ──
        {
            "type": "content",
            "slide_title": "音声を起点とした次世代業務自動化",
            "key_message": "「会議で話すだけで業務が動き出す」― LLMの進化がもたらす新しいワークフロー",
            "body_items": [
                {"heading": "Before（従来）", "body": "会議→メモ取り→手動でタスク登録→手動でレポート作成\n多くの手作業と時間ロスが発生"},
                {"heading": "After（Notta + LLM）", "body": "会議→Notta自動文字起こし→LLMが分析→タスク自動生成+レポート自動作成\n人手ゼロで業務が動く"}
            ],
            "notes": "前半の振り返り（セキュア基盤の実現）から、次のステップ「活用」へ。華パート開始。2分"
        },
        # ── スライド11: 技術アーキテクチャ ──
        {
            "type": "content",
            "slide_title": "Amazon Bedrock Agents × Notta で構築する業務自動化基盤",
            "key_message": "6層アーキテクチャ ― 音声入力からAI処理・実行・データ・接続・監視まで一気通貫",
            "body_items": [
                {"heading": "音声入力層", "body": "Notta（文字起こし + 要約）\n会議音声をテキスト化して次の層へ"},
                {"heading": "AI処理層", "body": "Amazon Bedrock Agents（Claude）\n自律的推論と多段階アクション実行\nKnowledge BasesによるRAG検索"},
                {"heading": "実行層", "body": "AWS Lambda: ツール関数（タスク生成/レポート作成等）\nAWS Step Functions: 多段階ワークフロー編排"},
                {"heading": "データ・接続・監視層", "body": "S3 / DynamoDB（データ層）\nAPI Gateway / SQS（接続層）\nCloudWatch（監視層）"}
            ],
            "notes": "6層構成のアーキテクチャ図をAWSサービスアイコンで表示。3分"
        },
        # ── スライド12: Bedrock Agents の仕組み ──
        {
            "type": "content",
            "slide_title": "Amazon Bedrock Agents ― 自律的なタスク実行の仕組み",
            "key_message": "ReAct パターンで指示を理解→計画立案→ツール実行→結果返却を自律的に実行",
            "body_items": [
                {"heading": "Agent指示", "body": "会議音声データから業務タスクを特定し自動実行する\nReAct（Reasoning + Acting）パターンによる自律的推論"},
                {"heading": "Action Groups", "body": "タスク生成ツール（Lambda）: ToDoリスト自動作成\nレポート生成ツール: 議事録→定型レポート\n通知ツール: Slack / メール自動通知"},
                {"heading": "Knowledge Base", "body": "過去の会議記録・社内ドキュメント\nRAG検索で関連情報を参照\nAgentの判断精度を向上"},
                {"heading": "セキュリティ", "body": "VPC内での閉域実行\nカレンダー登録ツール: 次回会議の自動スケジュール\nIAMによる最小権限アクセス"}
            ],
            "notes": "処理フロー図: Notta→文字起こし→Agent→推論→Lambda実行→結果統合。3分"
        },
        # ── スライド13: デモシナリオ ──
        {
            "type": "content",
            "slide_title": "ライブデモ ― 「会議から業務が自動で動き出す」",
            "key_message": "模擬会議の文字起こしから、タスク登録・議事録生成・Slack通知・次回日程提案まで自動実行",
            "body_items": [
                {"heading": "Step 1: 音声入力", "body": "Notta で模擬会議の音声を文字起こし"},
                {"heading": "Step 2: Agent分析", "body": "文字起こし結果をBedrock Agentに入力\nAgentが会議内容を自動分析"},
                {"heading": "Step 3: 自動実行", "body": "アクションアイテム抽出→タスク自動登録\n議事録サマリー自動生成\n関係者へSlack自動通知\n次回会議日程の提案"}
            ],
            "notes": "デモ画面のモックアップを表示。1分で説明後、ライブデモへ"
        },
        # ── スライド14: ライブデモ ──
        {
            "type": "content",
            "slide_title": "[ライブデモ] 音声起点の業務自動化Agent",
            "key_message": "Notta音声入力 → Bedrock Agent推論 → Lambda実行 ― セキュアVPC環境でのリアルタイムデモ",
            "body_items": [
                {"heading": "デモポイント", "body": "Notta音声入力の精度\nAgentの推論過程の可視化\n各ツール実行結果のリアルタイム表示\nセキュア環境（VPC内）での実行"},
                {"heading": "使用AWSサービス", "body": "Bedrock Agents / Lambda / Step Functions\nS3 / DynamoDB / API Gateway\nSQS / CloudWatch"}
            ],
            "notes": "5分のデモ実演。事前収録動画のバックアップあり。会場Wi-Fi＋モバイルテザリングをバックアップ回線として準備"
        },
        # ── スライド15: デモまとめ ──
        {
            "type": "content",
            "slide_title": "デモで見たアーキテクチャのポイント",
            "key_message": "セキュリティ・スケーラビリティ・拡張性 ― 3つの柱で実現するエンタープライズ業務自動化",
            "body_items": [
                {"heading": "セキュリティと利便性の両立", "body": "VPC + PrivateLink による閉域処理\nデータは日本国内AWSリージョンに保管\nAI学習なし設定でデータ非保持を実現"},
                {"heading": "スケーラビリティ", "body": "Lambda によるサーバーレス実行\nStep Functions による複雑なワークフロー管理\nSQS による非同期処理の安定性"},
                {"heading": "拡張性", "body": "新しいツール（Action Group）の追加が容易\nKnowledge Base の拡充でAgent精度向上\n外部SaaS連携（Salesforce, HubSpot等）への展開"}
            ],
            "notes": "3つのポイントを強調。2分"
        },
        # ── スライド16: ユースケース展開 ──
        {
            "type": "content",
            "slide_title": "業界別ユースケースと展開可能性",
            "key_message": "金融・官公庁・製造・医療 ― 各業界固有のセキュリティ要件に対応した音声DXを実現",
            "body_items": [
                {"heading": "金融機関", "body": "コンプライアンス対応の会議記録管理\n閉域網内での音声データ処理\n監査ログの自動保存"},
                {"heading": "官公庁・自治体", "body": "議会・審議会の議事録自動作成\nISMAP対応のセキュア環境\n多言語対応（外国人住民対応）"},
                {"heading": "製造業", "body": "設計レビュー会議の知識ベース化\n品質会議からの不具合管理自動化"},
                {"heading": "医療機関", "body": "HIPAA準拠の医療カンファレンス記録\n患者情報を含む音声データのセキュア処理"}
            ],
            "notes": "業界別4象限図。2分"
        },
        # ── スライド17: まとめ ──
        {
            "type": "content",
            "slide_title": "まとめ ― セキュアな音声基盤と業務自動化の融合",
            "key_message": "セキュア基盤 × LLM自動化 × エンタープライズ管理 ― Notta + AWS で音声DXの未来を実現",
            "body_items": [
                {"heading": "1. セキュア音声基盤", "body": "AWS VPC + PrivateLink + KMS で閉域音声処理を実現"},
                {"heading": "2. データ主権", "body": "日本国内DCでのデータ保管、AI学習なし設定でデータ非保持"},
                {"heading": "3. エンタープライズ管理", "body": "SSO / IP制限 / 監査ログによる多層防御"},
                {"heading": "4. 業務自動化", "body": "Bedrock Agents × Notta で「会議から業務が動く」を実現"},
                {"heading": "5. 拡張性", "body": "既存の業務ツール・SaaSとのシームレスな連携"}
            ],
            "notes": "5つのKey Takeawaysを1つずつ簡潔に振り返る。2分"
        },
        # ── スライド18: CTA・クロージング ──
        {
            "type": "content",
            "slide_title": "Notta で始めるセキュアな音声DX",
            "key_message": "エンタープライズプランで、セキュアな音声基盤と業務自動化をすぐに始められます",
            "body_items": [
                {"heading": "エンタープライズプラン", "body": "VPC接続 / SSO / IP制限 / 監査ログ\nすべてのセキュリティ機能をフル活用"},
                {"heading": "ブース・個別相談", "body": "展示ブースへぜひお越しください\n個別デモ・導入相談を承ります"},
                {"heading": "お問い合わせ", "body": "Notta公式サイト: notta.ai\n資料ダウンロード・お問い合わせフォーム\nご清聴ありがとうございました"}
            ],
            "notes": "QRコード表示。NottaロゴとQRコードを大きく配置。ブース番号を目立たせる。1分"
        },
    ]
}

# JSON保存
with open("aws_summit_2026_structure.json", "w", encoding="utf-8") as f:
    json.dump(presentation_data, f, ensure_ascii=False, indent=2)
print("JSON保存完了: aws_summit_2026_structure.json")

# PPTX生成
create_pptx(presentation_data, "aws_summit_2026_presentation.pptx")
print("完了！")
