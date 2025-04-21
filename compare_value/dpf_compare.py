from ansys.dpf import core as dpf
import os
import openpyxl
import time
import matplotlib.pyplot as plt
import pandas as pd


ROOT_PATH = os.getcwd()  # 현재 위치 가져오기
RESULT_PATH = f"{ROOT_PATH}/result"
COLUMNS = ["Time", "NodeID", "X", "Y", "Z", "Temperature"]


def isdir_and_make(dir_name: str):
    if not (os.path.isdir(dir_name)):
        os.makedirs(name=dir_name, exist_ok=True)
        print(f"Success: Create {dir_name}\n")


class DPF_Compare:
    def __init__(self, rst_path1, rst_path2, name1, name2):
        isdir_and_make(RESULT_PATH)  # result 폴더 생성
        self.rst_path = [rst_path1, rst_path2]
        self.name = [name1, name2]
        self.model = [None, None]
        self.disp_op = [None, None]
        self.temperature_op = [None, None]
        self.mesh = [None, None]
        self.metadata = [None, None]
        self.time_sec = [None, None]
        self.step_ids = [None, None]
        self.result = [None, None]
        self.node = []

        # 초기화 함수
        self._load_rst()  # model, rst 읽기

    def reload_model(self, input_rst_path1, input_rst_path2, input_name1, input_name2):
        self.rst_path = [input_rst_path1.strip(), input_rst_path2.strip()]
        self.name = [input_name1, input_name2]
        self._load_rst()  # model, rst 읽기
        print(f"설정된 경로: {self.rst_path}\n")

    def compare_x_disp(self, nodes):
        print("************************************************************")
        print("*** #1.Mode DPF Compare X-Disp and Temperature Start!\n"
              "*** Please Do not try additional command until it completes")
        print("************************************************************\n")
        exe_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
        print(f'Starting at: {exe_time}')
        isdir_and_make(f'{RESULT_PATH}/Mode 1')  # result 결과 시간 폴더 생성

        self.wrap_result(exe_time, nodes, 'Mode 1')
        self.draw_graph(exe_time, 'Mode 1')
        self.stop('#1.Mode')
        
    def compare_y_disp(self, nodes):
        print("************************************************************")
        print("*** #2.Mode DPF Compare Y-Disp and Temperature Start!\n"
              "*** Please Do not try additional command until it completes")
        print("************************************************************\n")
        exe_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
        print(f'Starting at: {exe_time}')
        isdir_and_make(f'{RESULT_PATH}/Mode 2')  # result 폴더 생성

        self.wrap_result(exe_time, nodes, 'Mode 2')
        self.draw_graph(exe_time, 'Mode 2')
        self.stop('#2.Mode')

    def compare_x_disp_with_nodes(self, nodes):
        print("************************************************************")
        print("*** #3.Mode DPF Compare X-Disp and Temperature with different nodes Start!\n"
              "*** Please Do not try additional command until it completes")
        print("************************************************************\n")
        exe_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
        print(f'Starting at: {exe_time}')
        isdir_and_make(f'{RESULT_PATH}/Mode 3')  # result 결과 시간 폴더 생성

        self.wrap_result_nodes(exe_time, nodes, 'Mode 3')
        self.draw_graph_nodes(exe_time, 'Mode 3')
        self.stop('#3.Mode')

    def compare_y_disp_with_nodes(self, nodes):
        print("************************************************************")
        print("*** #4.Mode DPF Compare Y-Disp and Temperature with different nodes Start!\n"
              "*** Please Do not try additional command until it completes")
        print("************************************************************\n")
        exe_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
        print(f'Starting at: {exe_time}')
        isdir_and_make(f'{RESULT_PATH}/Mode 4')  # result 결과 시간 폴더 생성

        self.wrap_result_nodes(exe_time, nodes, 'Mode 4')
        self.draw_graph_nodes(exe_time, 'Mode 4')
        self.stop('#4.Mode')

    def stop(self, mode: str):
        print(f'Ending at: {time.strftime("%a, %d-%b-%Y %I:%M:%S", time.localtime(time.time()))}')
        print("************************************************************")
        print(f"*** {mode} DPF Compare Disp and Temperature completed")
        print("************************************************************\n")

    def wrap_result(self, exe_time: str, node_ids: str, mode: str):
        isdir_and_make(f'{RESULT_PATH}/{mode}/{exe_time}')  # result 폴더 생성

        # 1. 파일 생성
        writer = pd.ExcelWriter(f'{RESULT_PATH}/{mode}/{exe_time}/result.xlsx', engine='openpyxl')

        self.node = [int(node.strip()) for node in node_ids.split(',')]
        mesh_scoping = dpf.mesh_scoping_factory.nodal_scoping(self.node)
        lst_result = []
        index = 1
        for model, meta in zip(self.model, self.metadata):
            time_freq = meta.time_freq_support.time_frequencies
            time_sec = time_freq.data
            step_ids = time_freq.scoping.ids
            time_index = 0
            lst_df_dis_temp = []
            for step in step_ids:
                time_steps_scoping = dpf.time_freq_scoping_factory.scoping_by_load_step(step)

                # Displacement
                disp_op = model.results.displacement()
                disp_op.inputs.mesh_scoping.connect(mesh_scoping)
                disp_op.inputs.time_scoping(time_steps_scoping)
                disp_out = disp_op.outputs.fields_container()

                # Temperature
                temperature_op = model.results.structural_temperature()
                temperature_op.inputs.requested_location.connect('Nodal')
                temperature_op.inputs.mesh_scoping.connect(mesh_scoping)
                temperature_op.inputs.time_scoping(time_steps_scoping)
                temperature_out = temperature_op.outputs.fields_container()

                # Result
                for dis, temp in zip(disp_out, temperature_out):
                    if float(time_sec[time_index]).is_integer():
                        df_dis = pd.DataFrame(data=dis.data, index=dis.scoping.ids, columns=['X', 'Y', 'Z']).sort_index(ascending=True)
                        df_temperature = pd.DataFrame(data=temp.data, index=temp.scoping.ids, columns=['Temperature']).sort_index(ascending=True)
                        df_dis_temp = pd.concat([df_dis, df_temperature], axis=1)
                        df_dis_temp['NodeID'] = df_dis_temp.index.values
                        df_dis_temp['Time'] = time_sec[time_index]
                        lst_df_dis_temp.append(df_dis_temp)
                    time_index += 1
            df_result = pd.concat(lst_df_dis_temp, axis=0).reset_index(drop=True)[COLUMNS]
            lst_result.append(df_result)
            df_result.to_excel(writer, sheet_name=f'model_{index}', index=False, header=True)
            index += 1
        self.result = lst_result

        # 3. 작성 완료 후 파일 저장
        writer.close()

    def draw_graph(self, exe_time: str, mode: str):
        axis = 'X' if mode == 'Mode 1' else 'Y'

        for node in self.node:
            plt.figure(figsize=(12, 6))

            # 각 컬럼에 대해 라인 플롯 생성
            x_max = 0
            for name, df_result in zip(self.name, self.result):
                df_node = df_result[df_result['NodeID'] == node]
                if x_max < df_node['Time'].max():
                    x_max = df_node['Time'].max()  # x축 최대값
                plt.plot(df_node['Time'], df_node[axis], '-o',
                         label=name,
                         linewidth=2,
                         markersize=4)

            # 그래프 꾸미기
            plt.title(f"{axis} Displacement vs Time, Node ID: {node}", fontsize=18, pad=15)
            plt.xlabel('Time(sec)', fontsize=12)
            # plt.xlim(0, x_max)  # x축 범위를 데이터의 최소값과 최대값으로 설정 x_data.min(), x_data.max()
            plt.ylabel(f'{axis} Displacement(mm)', fontsize=12)
            plt.legend(fontsize=10, bbox_to_anchor=(1.01, 1), loc='upper left')
            plt.grid(True, linestyle='--', alpha=0.7)

            # 여백 조정
            plt.tight_layout()
            plt.savefig(f'{RESULT_PATH}/{mode}/{exe_time}/Node_{node}_{axis}_disp.png', format='png')
            plt.cla()  # clear the current axes
            plt.clf()  # clear the current figure
            plt.close()  # closes the current figure

            plt.figure(figsize=(12, 6))

            # 각 컬럼에 대해 라인 플롯 생성
            x_max = 0
            for name, df_result in zip(self.name, self.result):
                df_node = df_result[df_result['NodeID'] == node]
                if x_max < df_node['Time'].max():
                    x_max = df_node['Time'].max()  # x축 최대값
                plt.plot(df_node['Time'], df_node['Temperature'], '-o',
                         label=name,
                         linewidth=2,
                         markersize=4)

            # 그래프 꾸미기
            plt.title(f"Temperature vs Time, Node ID: {node}", fontsize=18, pad=15)
            plt.xlabel('Time(sec)', fontsize=12)
            # plt.xlim(0, x_max)  # x축 범위를 데이터의 최소값과 최대값으로 설정 x_data.min(), x_data.max()
            plt.ylabel('Temperature(c)', fontsize=12)
            plt.legend(fontsize=10, bbox_to_anchor=(1.01, 1), loc='upper left')
            plt.grid(True, linestyle='--', alpha=0.7)

            # 여백 조정
            plt.tight_layout()
            plt.savefig(f'{RESULT_PATH}/{mode}/{exe_time}/Node_{node}_Temperature.png', format='png')
            plt.cla()  # clear the current axes
            plt.clf()  # clear the current figure
            plt.close()  # closes the current figure

    def wrap_result_nodes(self, exe_time: str, node_ids: str, mode: str):
        isdir_and_make(f'{RESULT_PATH}/{mode}/{exe_time}')  # result 폴더 생성

        # 1. 파일 생성
        writer = pd.ExcelWriter(f'{RESULT_PATH}/{mode}/{exe_time}/result.xlsx', engine='openpyxl')

        self.node = self._different_nodes(node_ids)

        lst_result = []
        index = 1
        for model, meta, node in zip(self.model, self.metadata, self.node):
            time_freq = meta.time_freq_support.time_frequencies
            time_sec = time_freq.data
            step_ids = time_freq.scoping.ids
            time_index = 0
            lst_df_dis_temp = []
            for step in step_ids:
                time_steps_scoping = dpf.time_freq_scoping_factory.scoping_by_load_step(step)
                mesh_scoping = dpf.mesh_scoping_factory.nodal_scoping(node)

                # Displacement
                disp_op = model.results.displacement()
                disp_op.inputs.mesh_scoping.connect(mesh_scoping)
                disp_op.inputs.time_scoping(time_steps_scoping)
                disp_out = disp_op.outputs.fields_container()

                # Temperature
                temperature_op = model.results.structural_temperature()
                temperature_op.inputs.requested_location.connect('Nodal')
                temperature_op.inputs.mesh_scoping.connect(mesh_scoping)
                temperature_op.inputs.time_scoping(time_steps_scoping)
                temperature_out = temperature_op.outputs.fields_container()

                # Result
                for dis, temp in zip(disp_out, temperature_out):
                    if float(time_sec[time_index]).is_integer():
                        df_dis = pd.DataFrame(data=dis.data, index=dis.scoping.ids, columns=['X', 'Y', 'Z']).sort_index(ascending=True)
                        df_temperature = pd.DataFrame(data=temp.data, index=temp.scoping.ids, columns=['Temperature']).sort_index(ascending=True)
                        df_dis_temp = pd.concat([df_dis, df_temperature], axis=1)
                        df_dis_temp['NodeID'] = df_dis_temp.index.values
                        df_dis_temp['Time'] = time_sec[time_index]
                        lst_df_dis_temp.append(df_dis_temp)
                    time_index += 1
            df_result = pd.concat(lst_df_dis_temp, axis=0).reset_index(drop=True)[COLUMNS]
            lst_result.append(df_result)
            df_result.to_excel(writer, sheet_name=f'model_{index}', index=False, header=True)
            index += 1
        self.result = lst_result

        # 3. 작성 완료 후 파일 저장
        writer.close()

    def draw_graph_nodes(self, exe_time: str, mode: str):
        axis = 'X' if mode == 'Mode 3' else 'Y'

        for node1, node2 in zip(self.node[0], self.node[1]):
            plt.figure(figsize=(12, 6))

            # 각 컬럼에 대해 라인 플롯 생성
            x_max = 0
            print(f"node1, node2: {node1, node2}")
            for name, df_result, node in zip(self.name, self.result, [node1, node2]):
                df_node = df_result[df_result['NodeID'] == node]
                if x_max < df_node['Time'].max():
                    x_max = df_node['Time'].max()  # x축 최대값
                plt.plot(df_node['Time'], df_node[axis], '-o',
                         label=name,
                         linewidth=2,
                         markersize=4)

            # 그래프 꾸미기
            plt.title(f"{axis} Displacement vs Time,  {self.name[0]} Node: {node1} vs {self.name[1]} Node: {node2}", fontsize=18, pad=15)
            plt.xlabel('Time(sec)', fontsize=12)
            # plt.xlim(0, x_max)  # x축 범위를 데이터의 최소값과 최대값으로 설정 x_data.min(), x_data.max()
            plt.ylabel(f'{axis} Displacement(mm)', fontsize=12)
            plt.legend(fontsize=10, bbox_to_anchor=(1.01, 1), loc='upper left')
            plt.grid(True, linestyle='--', alpha=0.7)

            # 여백 조정
            plt.tight_layout()
            plt.savefig(f'{RESULT_PATH}/{mode}/{exe_time}/Node_{node1}_{node2}_{axis}_disp.png', format='png')
            plt.cla()  # clear the current axes
            plt.clf()  # clear the current figure
            plt.close()  # closes the current figure

            plt.figure(figsize=(12, 6))

            # 각 컬럼에 대해 라인 플롯 생성
            x_max = 0
            for name, df_result, node in zip(self.name, self.result, [node1, node2]):
                df_node = df_result[df_result['NodeID'] == node]
                if x_max < df_node['Time'].max():
                    x_max = df_node['Time'].max()  # x축 최대값
                plt.plot(df_node['Time'], df_node['Temperature'], '-o',
                         label=name,
                         linewidth=2,
                         markersize=4)

            # 그래프 꾸미기
            plt.title(f"Temperature vs Time,  {self.name[0]} Node: {node1} vs {self.name[1]} Node: {node2}", fontsize=18, pad=15)
            plt.xlabel('Time(sec)', fontsize=12)
            # plt.xlim(0, x_max)  # x축 범위를 데이터의 최소값과 최대값으로 설정 x_data.min(), x_data.max()
            plt.ylabel('Temperature(c)', fontsize=12)
            plt.legend(fontsize=10, bbox_to_anchor=(1.01, 1), loc='upper left')
            plt.grid(True, linestyle='--', alpha=0.7)

            # 여백 조정
            plt.tight_layout()
            plt.savefig(f'{RESULT_PATH}/{mode}/{exe_time}/Node_{node1}_{node2}_Temperature.png', format='png')
            plt.cla()  # clear the current axes
            plt.clf()  # clear the current figure
            plt.close()  # closes the current figure

    def _load_rst(self):
        # rst 파일 찾기
        lst_model, lst_metadata, lst_mesh = [], [], []
        for i, rst_path in enumerate(self.rst_path):
            rst_file = [file for file in os.listdir(rst_path) if '.rst' in file]
            # rst 파일 직접 로드
            model = dpf.Model(os.path.join(rst_path, rst_file[0]))
            lst_model.append(model)
            lst_metadata.append(model.metadata)
            lst_mesh.append(model.metadata.meshed_region)

        self.model = lst_model
        self.metadata = lst_metadata
        self.mesh = lst_mesh

        for m in self.model:
            print(m)

    def _different_nodes(self, nodes_str: str):
        case1, case2 = [], []
        for node in nodes_str.split(')'):
            nodes = node.replace('(', '').split(', ')
            if len(nodes) > 1:
                case1.append(int(nodes[0].strip()))
                case2.append(int(nodes[1].strip()))
        return [case1, case2]
 