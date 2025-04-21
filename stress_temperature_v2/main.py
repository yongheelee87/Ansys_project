from dpf_stress_temperature import DPF_Stress_Temperature

RST_PATH = r"C:\Users\yongh\Ansys_HJ\Test\Test_files\dp0\SYS-7\MECH"  # 결과 파일이 있는 폴더
SAVE_RAW_DATA = False


if __name__ == "__main__":
    dpf_analysis = DPF_Stress_Temperature(save_data=SAVE_RAW_DATA, rst_path=RST_PATH)
    print('[INFO]\n'
          '#0.Mode: All Ansys Data to Parquet File\n'
          '#1.Mode: Run DPF Stress Temperature analysis\n'
          '#2.Mode: Run Temperature Range Check\n'
          '#3.Mode: Run DPF Stress Temperature analysis (all nodes)\n'
          '#4.Mode: Run DPF Stress Temperature analysis (specific nodes)\n'
          '#5.Mode: Run DPF Stress Temperature Static analysis\n'
          '#6.Mode: Run DPF Stress Temperature Transient analysis\n'
          '#Else. System Exit\n')
    mode = input('Please Press the Mode to Execute (모드 입력): ')
    while mode != '99':
        if mode == '0':
            dpf_analysis.all_data_to_parquet()
        elif mode == '1':
            dpf_analysis.analyze_base_model()
        elif mode == '2':
            dpf_analysis.check_temperature_range()
        elif mode == '3':
            dpf_analysis.analyze_base_model_all_nodes()
        elif mode == '4':
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.analyze_base_model_specific_nodes(node_ids)
        elif mode == '5':
            dpf_analysis.analyze_static()
        elif mode == '6':
            dpf_analysis.analyze_transient()
        else:
            print('잘못눌렀어요!! 정신차료\n')
            # sys.exit(0)

        print('[INFO]\n'
              '#0.Mode: All Ansys Data to Parquet File\n'
              '#1.Mode: Run DPF Stress Temperature analysis\n'
              '#2.Mode: Run Temperature Range Check\n'
              '#3.Mode: Run DPF Stress Temperature analysis (all nodes)\n'
              '#4.Mode: Run DPF Stress Temperature analysis (specific nodes)\n'
              '#5.Mode: Run DPF Stress Temperature Static analysis\n'
              '#6.Mode: Run DPF Stress Temperature Transient analysis\n'
              '#Else. System Exit\n')
        mode = input('Please Press the Mode to Execute (모드 입력): ')
