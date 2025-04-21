import pandas as pd
import matplotlib.pyplot as plt
import os
import time

RESULT_PATH = './result'


def isdir_and_make(dir_name: str):
    if not (os.path.isdir(dir_name)):
        os.makedirs(name=dir_name, exist_ok=True)
        print(f"Success: Create {dir_name}\n")


if __name__ == "__main__":
    isdir_and_make(RESULT_PATH)
    # 엑셀 파일 읽기
    df = pd.read_excel('displacement.xlsx')
    unit = df.iloc[1, 2]
    df = df.iloc[2:, :]
    lst_element = sorted(list(set([col.split()[0] for col in df.columns.tolist()[1:]])))
    pairs = []
    for ele in lst_element:
        pair = [col for col in df.columns if ele in col]
        pairs.append(pair)

    exe_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
    isdir_and_make(f"{RESULT_PATH}/{exe_time}")
    for pair in pairs:
        x_col = df.columns[0]
        # 그래프 생성
        fig, ax = plt.subplots(figsize=(10, 6))

        # 첫 번째 컬럼 플롯
        ax.plot(df[x_col], df[pair[0]], marker='o', linestyle='-', linewidth=2, label=pair[0])

        # 두 번째 컬럼 플롯
        ax.plot(df[x_col], df[pair[1]], marker='s', linestyle='--', linewidth=2, label=pair[1])

        # 그래프 꾸미기
        ax.set_title(f'{pair[0]} vs {pair[1]}', fontsize=16)

        ax.set_xlabel('Time', fontsize=12)
        ax.set_ylabel(unit, fontsize=12)
        ax.legend(fontsize=12, loc='upper right')
        ax.grid(True, alpha=0.3)

        # x축 눈금 조정 (데이터 포인트가 많으면 적절히 조정)
        if len(df) > 20:
            plt.xticks(rotation=45)
            # 모든 x 값을 표시하지 않고 적절히 간격 조정
            ax.xaxis.set_major_locator(plt.MaxNLocator(10))

        plt.tight_layout()

        plt.savefig(f'{RESULT_PATH}/{exe_time}/{pair[0]} vs {pair[1]}.png', dpi=300)
