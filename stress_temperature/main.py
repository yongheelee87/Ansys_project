from dpf_stress_temperature import DPF_Stress_Temperature

RST_PATH = r"C:\Users\yongh\Ansys_HJ\Test\Test_files\dp0\SYS-7\MECH"  # 결과 파일이 있는 폴더
SAVE_RAW_DATA = True


if __name__ == "__main__":
    dpf_analysis = DPF_Stress_Temperature(save_data=SAVE_RAW_DATA, rst_path=RST_PATH)
    print('[INFO]\n'
          '#0.Mode: Reload Model and RST path\n'
          '#1.Mode: Run DPF Stress Temperature analysis\n'
          '#2.Mode: Run Temperature Range Check\n'
          '#3.Mode: Run DPF Stress Temperature analysis (all nodes)\n'
          '#4.Mode: Run DPF Stress Temperature analysis (specific nodes)\n'
          '#Else. System Exit\n')
    mode = input('Please Press the Mode to Execute (모드 입력): ')
    while mode != '99':
        if mode == '0':
            rst_path = input('RST 파일 경로를 입력해 주세요: ')
            dpf_analysis.reload_model(input_rst_path=rst_path.strip())
        elif mode == '1':
            dpf_analysis.analyze_base_model()
        elif mode == '2':
            dpf_analysis.check_temperature_range()
        elif mode == '3':
            dpf_analysis.analyze_base_model_all_nodes()
        elif mode == '4':
            material = input('소재 약어를 입력해 주세요: ')
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.analyze_base_model_specific_nodes(material, node_ids)
        elif mode == '7':
            dpf_analysis.out_all_nodes_raw()
        else:
            print('잘못눌렀어요!! 정신차료\n')
            # sys.exit(0)

        print('[INFO]\n'
              '#0.Mode: Reload Model and RST path\n'
              '#1.Mode: Run DPF Stress Temperature analysis\n'
              '#2.Mode: Run Temperature Range Check\n'
              '#3.Mode: Run DPF Stress Temperature analysis (all nodes)\n'
              '#4.Mode: Run DPF Stress Temperature analysis (specific nodes)\n'
              '#Else. System Exit\n')
        mode = input('Please Press the Mode to Execute (모드 입력): ')
