import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 파일 경로 설정
file_path = './data/2.elevator_failure_prediction (2).xlsx'
results_dir = './data/results'
results_file = os.path.join(results_dir, 'results.xlsx')

# 결과 디렉토리 생성
if not os.path.exists(results_dir):
    os.makedirs(results_dir)

try:
    # Excel 파일 읽기
    data = pd.ExcelFile(file_path)
    print("파일을 성공적으로 읽었습니다!")
    print("시트 목록:", data.sheet_names)  # 파일에 포함된 시트 이름 출력

    # 특정 시트 읽기 예시 (기본 시트 읽기)
    df = data.parse(sheet_name=0)  # 첫 번째 시트를 DataFrame으로 읽기
    print("데이터 프레임 미리보기:")
    print(df.head())  # 데이터의 처음 몇 줄 출력

    # 데이터 분석: 각 열의 데이터 타입 확인
    print("데이터 타입 정보:")
    print(df.dtypes)

    # 결측치 확인
    print("결측치 개수:")
    print(df.isnull().sum())

    # 결측치 처리 (예: 결측치를 0으로 채우기)
    df.fillna(0, inplace=True)
    print("결측치 처리 후 데이터 프레임 미리보기:")
    print(df.head())

    # 주요 열에 대한 데이터 분석
    important_columns = ["Time", "Temperature", "Humidity", "RPM", "Vibrations", "Pressure",
                         "Sensor1", "Sensor2", "Sensor3", "Sensor4", "Sensor5", "Sensor6", "Status"]

    # 주요 열만 선택
    df = df[important_columns]
    print("주요 열 선택 후 데이터 미리보기:")
    print(df.head())

    # Status 분포 확인
    plt.figure()
    sns.countplot(x="Status", data=df)
    plt.title("Status Distribution")
    plt.xlabel("Status")
    plt.ylabel("Count")
    plt.savefig(os.path.join(results_dir, 'status_distribution.png'))
    plt.close()

    # 각 센서 데이터의 분포 확인
    sensor_columns = ["Sensor1", "Sensor2", "Sensor3", "Sensor4", "Sensor5", "Sensor6"]
    for sensor in sensor_columns:
        plt.figure()
        sns.histplot(df[sensor], kde=True)
        plt.title(f"Distribution of {sensor}")
        plt.xlabel(sensor)
        plt.ylabel("Frequency")
        plt.savefig(os.path.join(results_dir, f'{sensor}_distribution.png'))
        plt.close()

    # 상관 관계 히트맵
    plt.figure(figsize=(12, 10))
    correlation_matrix = df.corr()
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt='.2f')
    plt.title("Correlation Matrix of Features")
    plt.savefig(os.path.join(results_dir, 'correlation_matrix.png'))
    plt.close()

    # Status에 따른 주요 변수 평균값 비교
    status_means = df.groupby("Status").mean()
    print("Status별 주요 변수 평균값:")
    print(status_means)

    # Status별 센서 데이터 시각화 (Boxplot)
    for sensor in sensor_columns:
        plt.figure()
        sns.boxplot(x="Status", y=sensor, data=df)
        plt.title(f"{sensor} by Status")
        plt.xlabel("Status")
        plt.ylabel(sensor)
        plt.savefig(os.path.join(results_dir, f'{sensor}_boxplot.png'))
        plt.close()

    # 결과를 Excel 파일로 저장
    with pd.ExcelWriter(results_file) as writer:
        df.to_excel(writer, sheet_name="Processed Data", index=False)
        status_means.to_excel(writer, sheet_name="Status Means")
    print(f"결과가 '{results_dir}'에 저장되었습니다.")

except FileNotFoundError:
    print("파일을 찾을 수 없습니다. 업로드된 경로가 정확한지 확인하거나 파일 경로를 수정하세요.")
    print("현재 경로:", file_path)

except PermissionError:
    print("파일에 대한 권한이 없습니다. 권한을 확인하거나 관리자 권한으로 실행하세요.")

except ValueError as ve:
    print("파일 형식이 잘못되었거나 지원되지 않는 형식입니다. 자세한 내용:", ve)

except Exception as e:
    print(f"예기치 못한 오류가 발생했습니다: {e}")