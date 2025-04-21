from dpf_compare import DPF_Compare

RST_PATH_1 = r"C:\Users\yongh\Ansys_HJ\Test\Test_files\dp0\SYS-3\MECH"  # 결과 파일이 있는 폴더
RST_PATH_2 = r"C:\Users\yongh\Ansys_HJ\Test\Test_files\dp0\SYS-2\MECH"  # 결과 파일이 있는 폴더
NAME_1 = "Case1"
NAME_2 = "Case2"


if __name__ == "__main__":
    dpf_analysis = DPF_Compare(rst_path1=RST_PATH_1, rst_path2=RST_PATH_2, name1=NAME_1, name2=NAME_2)
    print('[INFO]\n'
          '#0.Mode: Reload Model and RST path\n'
          '#1.Mode: Run Compare X_displacement and Temperature\n'
          '#2.Mode: Run Compare Y_displacement and Temperature\n'
          '#3.Mode: Run Compare X_displacement and Temperature with different nodes\n'
          '#4.Mode: Run Compare Y_displacement and Temperature with different nodes\n'
          '#Else. System Exit\n')
    mode = input('Please Press the Mode to Execute (모드 입력): ')
    while mode != '99':
        if mode == '0':
            rst_path_1 = input('첫번째 RST 파일 경로를 입력해 주세요: ')
            name_1 = input('첫번째 케이스를 정의해주세요: ')
            rst_path_2 = input('두번째 RST 파일 경로를 입력해 주세요: ')
            name_2 = input('두번째 케이스를 정의해주세요: ')
            dpf_analysis.reload_model(input_rst_path1=rst_path_1, input_rst_path2=rst_path_2, input_name1=name_1, input_name2=name_2)
        elif mode == '1':
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.compare_x_disp(node_ids)
        elif mode == '2':
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.compare_y_disp(node_ids)
        elif mode == '3':
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.compare_x_disp_with_nodes(node_ids)
        elif mode == '4':
            node_ids = input('노드 ID를 입력해 주세요: ')
            dpf_analysis.compare_y_disp_with_nodes(node_ids)
        else:
            print('잘못눌렀어요!! 정신차료\n')
            # sys.exit(0)

        print('[INFO]\n'
              '#0.Mode: Reload Model and RST path\n'
              '#1.Mode: Run Compare X_displacement and Temperature\n'
              '#2.Mode: Run Compare Y_displacement and Temperature\n'
              '#3.Mode: Run Compare X_displacement and Temperature with different nodes\n'
              '#4.Mode: Run Compare Y_displacement and Temperature with different nodes\n'
              '#Else. System Exit\n')
        mode = input('Please Press the Mode to Execute (모드 입력): ')
