import subprocess
import sys

# インストールするパッケージと推奨されるバージョン
# 依存関係を考慮して順序を調整しています
packages = {
    # 1. ツールと基本ライブラリ
    "setuptools": "57.4.0",
    "pip": "25.2", 
    "six": "1.17.0",
    "numpy": "2.0.2",
    "typing_extensions": "4.15.0",
    "packaging": "25.0",
    "zipp": "3.23.0",
    "importlib_metadata": "8.7.0",
    
    # 2. ortoolsとその依存関係
    "absl-py": "2.3.1",
    "protobuf": "3.20.3",
    "immutabledict": "4.2.2",
    "ortools": "9.0.9048",
    
    # 3. pandasとその依存関係
    "python-dateutil": "2.9.0.post0",
    "pytz": "2025.2",
    "tzdata": "2025.2",
    "pandas": "2.3.3",
    
    # 4. openpyxlとその依存関係
    "et_xmlfile": "2.0.0",
    "openpyxl": "3.1.5",
    
    # 5. pyinstallerとその依存関係
    "altgraph": "0.17.4",
    "pefile": "2023.2.7",
    "pywin32-ctypes": "0.2.3",
    "pyinstaller-hooks-contrib": "2025.9",
    "pyinstaller": "6.16.0",
}

def install_package(package_name, version):
    """指定されたパッケージとバージョンをインストールします。--trusted-hostオプションを含みます。"""
    try:
        # pip install パッケージ名==バージョン --trusted-host pypi.org --trusted-host files.pythonhosted.org
        command = [
            sys.executable,
            "-m", 
            "pip", 
            "install", 
            f"{package_name}=={version}",
            "--trusted-host", "pypi.org",
            "--trusted-host", "files.pythonhosted.org"
        ]
        
        print(f"\n--- インストール開始: **{package_name}** ({version}) ---")
        
        # コマンドを実行し、出力をリアルタイムで表示
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        for line in process.stdout:
            print(line, end="")
        process.wait()
        
        if process.returncode == 0:
            print(f"✅ {package_name} のインストールに成功しました。")
        else:
            print(f"❌ {package_name} のインストールに失敗しました (終了コード: {process.returncode})。")
            return False
    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")
        return False
    return True

# メイン処理
def main():
    print("--- 🔒 **TRUSTED-HOST**オプションを使用してパッケージインストールを開始します ---")
    print("⚠ 注意: この方法はSSL検証をスキップするため、セキュリティリスクを伴います。")
    
    for package, version in packages.items():
        if not install_package(package, version):
            print(f"\n🚨 **エラーが発生したため、インストールを中止します。**")
            break
    else:
        print("\n✨ **すべてのパッケージのインストールが完了しました。**")

if __name__ == "__main__":
    main()