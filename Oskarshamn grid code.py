import pandapower as pp
import numpy as np
import xlsxwriter as xl
import pandapower.plotting as pplt
import matplotlib.pyplot as plt
import pandas as pd
import os
import tempfile
from pandapower.timeseries import DFData
from pandapower.timeseries import OutputWriter
from pandapower.timeseries.run_time_series import run_timeseries
from pandapower.control import ConstControl

#df = pd.read_excel('CON_values_2.xlsx', skiprows=1)

#create empty net
net = pp.create_empty_network()

#buses
b0 = pp.create_bus(net, vn_kv=50., name="Bus 0")
b1_virtual = pp.create_bus(net, vn_kv=50, name="Bus 1")
b2 = pp.create_bus(net, vn_kv=12, name="Bus 2")
b3 = pp.create_bus(net, vn_kv=0.4, name="Bus 3")
b4 = pp.create_bus(net, vn_kv=12, name="Bus 4")
b5 = pp.create_bus(net, vn_kv=0.4, name="Bus 5")
b6 = pp.create_bus(net, vn_kv=12, name="Bus 6")
b7 = pp.create_bus(net, vn_kv=0.4, name="Bus 7")
b8 = pp.create_bus(net, vn_kv=12, name="Bus 8")
b9 = pp.create_bus(net, vn_kv=0.4, name="Bus 9")
b10_virtual = pp.create_bus(net, vn_kv=12, name="Bus 10")
b11_virtual = pp.create_bus(net, vn_kv=12, name="Bus 11")
b12_virtual = pp.create_bus(net, vn_kv=12, name="Bus 12")
b13 = pp.create_bus(net, vn_kv=12, name="Bus 13")
b14 = pp.create_bus(net, vn_kv=0.4, name="Bus 14")
b15_virtual = pp.create_bus(net, vn_kv=12, name="Bus 15")
b16_virtual = pp.create_bus(net, vn_kv=12, name="Bus 16")
b17_virtual = pp.create_bus(net, vn_kv=12, name="Bus 17")
b18_virtual = pp.create_bus(net, vn_kv=12, name="Bus 18")
b19 = pp.create_bus(net, vn_kv=12, name="Bus 19")
b20 = pp.create_bus(net, vn_kv=0.4, name="Bus 20")
b21_virtual = pp.create_bus(net, vn_kv=12, name="Bus 21")
b22_virtual = pp.create_bus(net, vn_kv=12, name="Bus 22")
b23 = pp.create_bus(net, vn_kv=12, name="Bus 23")
b24 = pp.create_bus(net, vn_kv=0.4, name="Bus 24")
b25 = pp.create_bus(net, vn_kv=12, name="Bus 25")
b26 = pp.create_bus(net, vn_kv=0.4, name="Bus 26")
b27_virtual = pp.create_bus(net, vn_kv=12, name="Bus 27")
b28 = pp.create_bus(net, vn_kv=12, name="Bus 28")
b29 = pp.create_bus(net, vn_kv=0.4, name="Bus 29")
b30_virtual = pp.create_bus(net, vn_kv=12, name="Bus 30")
b31_virtual = pp.create_bus(net, vn_kv=12, name="Bus 31")
b32_virtual = pp.create_bus(net, vn_kv=12, name="Bus 32")
b33 = pp.create_bus(net, vn_kv=12, name="Bus 33")
b34 = pp.create_bus(net, vn_kv=0.4, name="Bus 34")
b35_virtual = pp.create_bus(net, vn_kv=12, name="Bus 35")
b36 = pp.create_bus(net, vn_kv=12, name="Bus 36")
b37 = pp.create_bus(net, vn_kv=0.4, name="Bus 37")
b38 = pp.create_bus(net, vn_kv=12, name="Bus 38")
b39 = pp.create_bus(net, vn_kv=0.4, name="Bus 39")
b40 = pp.create_bus(net, vn_kv=12, name="Bus 40")
b41 = pp.create_bus(net, vn_kv=0.4, name="Bus 41")
b42_virtual = pp.create_bus(net, vn_kv=12, name="Bus 42")
b43_virtual = pp.create_bus(net, vn_kv=12, name="Bus 43")
b44 = pp.create_bus(net, vn_kv=12, name="Bus 44")
b45 = pp.create_bus(net, vn_kv=0.4, name="Bus 45")
b46 = pp.create_bus(net, vn_kv=12, name="Bus 46")
b47 = pp.create_bus(net, vn_kv=0.4, name="Bus 47")
b48 = pp.create_bus(net, vn_kv=12, name="Bus 48")
b49 = pp.create_bus(net, vn_kv=0.4, name="Bus 49")
b50_virtual = pp.create_bus(net, vn_kv=12, name="Bus 50")
b51 = pp.create_bus(net, vn_kv=12, name="Bus 51")
b52 = pp.create_bus(net, vn_kv=0.4, name="Bus 52")
b53 = pp.create_bus(net, vn_kv=12, name="Bus 53")
b54 = pp.create_bus(net, vn_kv=0.4, name="Bus 54")
b55 = pp.create_bus(net, vn_kv=12, name="Bus 55")
b56 = pp.create_bus(net, vn_kv=0.4, name="Bus 56")
b57 = pp.create_bus(net, vn_kv=12, name="Bus 57")
b58 = pp.create_bus(net, vn_kv=0.4, name="Bus 58")


#bus elements
pp.create_ext_grid(net, bus=b0, vm_pu=1.0, name="ALVARSBERG")
pp.create_load(net, bus=b0, p_mw=0, q_mvar=0, name="Bus 0 Load")
pp.create_load(net, bus=b1_virtual, p_mw=0, q_mvar=0, name="Bus 1 Load")
pp.create_load(net, bus=b2, p_mw=0, q_mvar=0, name="Bus 2 Load")
pp.create_load(net, bus=b3, p_mw=0, q_mvar=0, name="Bus 3 Load")
pp.create_load(net, bus=b4, p_mw=0, q_mvar=0, name="Bus 4 Load")
pp.create_load(net, bus=b5, p_mw=0, q_mvar=0, name="Bus 5 Load")
pp.create_load(net, bus=b6, p_mw=0, q_mvar=0, name="Bus 6 Load")
pp.create_load(net, bus=b7, p_mw=0, q_mvar=0, name="Bus 7 Load")
pp.create_load(net, bus=b8, p_mw=0, q_mvar=0, name="Bus 8 Load")
pp.create_load(net, bus=b9, p_mw=0, q_mvar=0, name="Bus 9 Load")
pp.create_load(net, bus=b10_virtual, p_mw=0, q_mvar=0, name="Bus 10 Load")
pp.create_load(net, bus=b11_virtual, p_mw=0, q_mvar=0, name="Bus 11 Load")
pp.create_load(net, bus=b12_virtual, p_mw=0, q_mvar=0, name="Bus 12 Load")
pp.create_load(net, bus=b13, p_mw=0, q_mvar=0, name="Bus 13 Load")
pp.create_load(net, bus=b14, p_mw=0, q_mvar=0, name="Load LAX SMHR")
pp.create_load(net, bus=b15_virtual, p_mw=0, q_mvar=0, name="Bus 15 Load")
pp.create_load(net, bus=b16_virtual, p_mw=0, q_mvar=0, name="Bus 16 Load")
pp.create_load(net, bus=b17_virtual, p_mw=0, q_mvar=0, name="Bus 17 Load")
pp.create_load(net, bus=b18_virtual, p_mw=0, q_mvar=0, name="Bus 18 Load")
pp.create_load(net, bus=b19, p_mw=0, q_mvar=0, name="Bus 19 Load")
pp.create_load(net, bus=b20, p_mw=0, q_mvar=0, name="Bus 20 Load")
pp.create_load(net, bus=b21_virtual, p_mw=0, q_mvar=0, name="Bus 21 Load")
pp.create_load(net, bus=b22_virtual, p_mw=0, q_mvar=0, name="Bus 22 Load")
pp.create_load(net, bus=b23, p_mw=0, q_mvar=0, name="Bus 23 Load")
pp.create_load(net, bus=b24, p_mw=0, q_mvar=0, name="Bus 24 Load")
pp.create_load(net, bus=b25, p_mw=0, q_mvar=0, name="Bus 25 Load")
pp.create_load(net, bus=b26, p_mw=0, q_mvar=0, name="Bus 26 Load")
pp.create_load(net, bus=b27_virtual, p_mw=0, q_mvar=0, name="Bus 27 Load")
pp.create_load(net, bus=b28, p_mw=0, q_mvar=0, name="Bus 28 Load")
pp.create_load(net, bus=b29, p_mw=0, q_mvar=0, name="Load BVH SMHR")
pp.create_load(net, bus=b30_virtual, p_mw=0, q_mvar=0, name="Bus 30 Load")
pp.create_load(net, bus=b31_virtual, p_mw=0, q_mvar=0, name="Bus 31 Load")
pp.create_load(net, bus=b32_virtual, p_mw=0, q_mvar=0, name="Bus 32 Load")
pp.create_load(net, bus=b33, p_mw=0, q_mvar=0, name="Bus 33 Load")
pp.create_load(net, bus=b34, p_mw=0, q_mvar=0, name="Load KLU SMHR")
pp.create_load(net, bus=b35_virtual, p_mw=0, q_mvar=0, name="Bus 35 Load")
pp.create_load(net, bus=b36, p_mw=0, q_mvar=0, name="Bus 36 Load")
pp.create_load(net, bus=b37, p_mw=0, q_mvar=0, name="Bus 37 Load")
pp.create_load(net, bus=b38, p_mw=0, q_mvar=0, name="Bus 38 Load")
pp.create_load(net, bus=b39, p_mw=0, q_mvar=0, name="Load OLJ SMHR")
pp.create_load(net, bus=b40, p_mw=0, q_mvar=0, name="Bus 40 Load")
pp.create_load(net, bus=b41, p_mw=0, q_mvar=0, name="Bus 41 Load")
pp.create_load(net, bus=b42_virtual, p_mw=0, q_mvar=0, name="Bus 42 Load")
pp.create_load(net, bus=b43_virtual, p_mw=0, q_mvar=0, name="Bus 43 Load")
pp.create_load(net, bus=b44, p_mw=0, q_mvar=0, name="Bus 44 Load")
pp.create_load(net, bus=b45, p_mw=0, q_mvar=0, name="Bus 45 Load")
pp.create_load(net, bus=b46, p_mw=0, q_mvar=0, name="Bus 46 Load")
pp.create_load(net, bus=b47, p_mw=0, q_mvar=0, name="Bus 47 Load")
pp.create_load(net, bus=b48, p_mw=0, q_mvar=0, name="Bus 48 Load")
pp.create_load(net, bus=b49, p_mw=0, q_mvar=0, name="Bus 49 Load")
pp.create_load(net, bus=b50_virtual, p_mw=0, q_mvar=0, name="Bus 50 Load")
pp.create_load(net, bus=b51, p_mw=0, q_mvar=0, name="Bus 51 Load")
pp.create_load(net, bus=b52, p_mw=0, q_mvar=0, name="Bus 52 Load")
pp.create_load(net, bus=b53, p_mw=0, q_mvar=0, name="Bus 53 Load")
pp.create_load(net, bus=b54, p_mw=0, q_mvar=0, name="Bus 54 Load")
pp.create_load(net, bus=b55, p_mw=0, q_mvar=0, name="Bus 55 Load")
pp.create_load(net, bus=b56, p_mw=0, q_mvar=0, name="Bus 56 Load")
pp.create_load(net, bus=b57, p_mw=0, q_mvar=0, name="Bus 57 Load")
pp.create_load(net, bus=b58, p_mw=0, q_mvar=0, name="Load JUN SMHR")



#branch elements
#tranformer paremeters

trafo_1_DGP = pp.create_transformer_from_parameters(net, hv_bus=b2, lv_bus=b3, sn_mva=0.5, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="OGP")
trafo_2_TBD = pp.create_transformer_from_parameters(net, hv_bus=b4, lv_bus=b5, sn_mva=0.5, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="TBD")
trafo_3_URA = pp.create_transformer_from_parameters(net, hv_bus=b6, lv_bus=b7, sn_mva=0.5, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="URA")
trafo_4_SVN = pp.create_transformer_from_parameters(net, hv_bus=b8, lv_bus=b9, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="SVN")
trafo_5_LAX = pp.create_transformer_from_parameters(net, hv_bus=b13, lv_bus=b14, sn_mva=1.6, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="LAX")
trafo_6_KBV = pp.create_transformer_from_parameters(net, hv_bus=b19, lv_bus=b20, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KBV")
trafo_7_NSG = pp.create_transformer_from_parameters(net, hv_bus=b23, lv_bus=b24, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="NSG")
trafo_8_KRVA = pp.create_transformer_from_parameters(net, hv_bus=b25, lv_bus=b26, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KRVA")
trafo_9_1 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN1")
trafo_9_2 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN2")
trafo_9_3 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=1.0, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN3")
trafo_9_4 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN4")
trafo_9_5 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=1.0, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN5")
trafo_9_6 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=1.0, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN6")
trafo_9_7 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN7")
trafo_9_8 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN8")
trafo_9_9 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=1.0, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN9")
trafo_9_10 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=1.0, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN10")
trafo_9_11 = pp.create_transformer_from_parameters(net, hv_bus=b57, lv_bus=b58, sn_mva=0.5, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="JUN11")
trafo_10 = pp.create_transformer_from_parameters(net, hv_bus=b28, lv_bus=b29, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="BVH")
trafo_11 = pp.create_transformer_from_parameters(net, hv_bus=b33, lv_bus=b34, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KLU")
trafo_12 = pp.create_transformer_from_parameters(net, hv_bus=b36, lv_bus=b37, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="SPL")
trafo_13 = pp.create_transformer_from_parameters(net, hv_bus=b38, lv_bus=b39, sn_mva=1.3, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="OLJ")
trafo_14 = pp.create_transformer_from_parameters(net, hv_bus=b40, lv_bus=b41, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="HDM")
trafo_15 = pp.create_transformer_from_parameters(net, hv_bus=b44, lv_bus=b45, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KAP")
trafo_16 = pp.create_transformer_from_parameters(net, hv_bus=b46, lv_bus=b47, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="HUK")
trafo_17 = pp.create_transformer_from_parameters(net, hv_bus=b48, lv_bus=b49, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="LAN")
trafo_18 = pp.create_transformer_from_parameters(net, hv_bus=b51, lv_bus=b52, sn_mva=0.5, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="MAN")
trafo_19 = pp.create_transformer_from_parameters(net, hv_bus=b53, lv_bus=b54, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KSO")
trafo_20 = pp.create_transformer_from_parameters(net, hv_bus=b55, lv_bus=b56, sn_mva=0.8, vn_hv_kv=12, vn_lv_kv=0.4,
                                                vk_percent=12, vkr_percent=0.41, pfe_kw=14, i0_percent=0.07,name="KOL")

pp.create_line(net, from_bus=b0, to_bus=b1_virtual, length_km=0.329, name="Line 0.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b1_virtual, to_bus=b2, length_km=0.046, name="Line 0.2",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b2, to_bus=b4, length_km=0.081, name="Line 1",std_type="679-AL1/86-ST1A 380.0")  

pp.create_line(net, from_bus=b4, to_bus=b6, length_km=0.334, name="Line 2",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b6, to_bus=b8, length_km=0.660, name="Line 3",std_type="679-AL1/86-ST1A 380.0")  

pp.create_line(net, from_bus=b8, to_bus=b10_virtual, length_km=0.453, name="Line 4.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b10_virtual, to_bus=b11_virtual, length_km=0.123, name="Line 4.2",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b11_virtual, to_bus=b12_virtual, length_km=0.136, name="Line 4.3",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b12_virtual, to_bus=b13, length_km=0.299, name="Line 4.4",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b13, to_bus=b15_virtual, length_km=0.063, name="Line 5.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b15_virtual, to_bus=b16_virtual, length_km=0.118, name="Line 5.2",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b16_virtual, to_bus=b17_virtual, length_km=0.021, name="Line 5.3",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b17_virtual, to_bus=b18_virtual, length_km=0.079, name="Line 5.4",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b18_virtual, to_bus=b19, length_km=0.112, name="Line 5.5",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b19, to_bus=b21_virtual, length_km=0.067, name="Line 6.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b21_virtual, to_bus=b22_virtual, length_km=0.282, name="Line 6.2",std_type="679-AL1/86-ST1A 380.0")  
pp.create_line(net, from_bus=b22_virtual, to_bus=b23, length_km=0.063, name="Line 6.3",std_type="679-AL1/86-ST1A 380.0")  

pp.create_line(net, from_bus=b23, to_bus=b25, length_km=0.143, name="Line 7", std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b23, to_bus=b57, length_km=0.497, name="Line 19", std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b19, to_bus=b27_virtual, length_km=0.181, name="Line 8.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b27_virtual, to_bus=b28, length_km=0.116, name="Line 8.2",std_type="679-AL1/86-ST1A 380.0")  

pp.create_line(net, from_bus=b28, to_bus=b30_virtual, length_km=0.288, name="Line 9.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b30_virtual, to_bus=b31_virtual, length_km=0.118, name="Line 9.2",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b31_virtual, to_bus=b32_virtual, length_km=0.185, name="Line 9.3",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b32_virtual, to_bus=b33, length_km=0.185, name="Line 9.4",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b33, to_bus=b35_virtual, length_km=0.265, name="Line 10.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b35_virtual, to_bus=b36, length_km=0.293, name="Line 10.2",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b36, to_bus=b38, length_km=0.315, name="Line 11",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b38, to_bus=b40, length_km=0.441, name="Line 12",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b38, to_bus=b42_virtual, length_km=0.251, name="Line 13.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b42_virtual, to_bus=b43_virtual, length_km=0.139, name="Line 13.2",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b43_virtual, to_bus=b44, length_km=0.215, name="Line 13.3",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b44, to_bus=b46, length_km=0.729, name="Line 14",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b46, to_bus=b48, length_km=0.417, name="Line 15",std_type="679-AL1/86-ST1A 380.0")  

pp.create_line(net, from_bus=b48, to_bus=b50_virtual, length_km=0.473, name="Line 16.1",std_type="679-AL1/86-ST1A 380.0")
pp.create_line(net, from_bus=b50_virtual, to_bus=b51, length_km=0.173, name="Line 16.2",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b44, to_bus=b53, length_km=0.398, name="Line 17",std_type="679-AL1/86-ST1A 380.0")

pp.create_line(net, from_bus=b53, to_bus=b55, length_km=0.366, name="Line 18",std_type="679-AL1/86-ST1A 380.0")  

#line diameters
net.line.loc[net.line['name'] == 'Line 0.1', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 0.2', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 1', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 2', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 3', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 4.1', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 4.2', 'diameter'] = 0.720
net.line.loc[net.line['name'] == 'Line 4.3', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 4.4', 'diameter'] = 0.098
net.line.loc[net.line['name'] == 'Line 5.1', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 5.2', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 5.3', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 5.4', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 5.5', 'diameter'] = 0.720
net.line.loc[net.line['name'] == 'Line 6.1', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 6.2', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 6.3', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 7', 'diameter'] = 0.720
net.line.loc[net.line['name'] == 'Line 8.1', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 8.2', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 9.1', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 9.2', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 9.3', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 9.4', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 10.1', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 10.2', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 11', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 12', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 13.1', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 13.2', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 13.3', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 14', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 15', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 16.1', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 16.2', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 17', 'diameter'] = 0.285
net.line.loc[net.line['name'] == 'Line 18', 'diameter'] = 0.450
net.line.loc[net.line['name'] == 'Line 19', 'diameter'] = 0.720

profilesA = pd.read_excel('Net_Bus_Present_Max_Pmw_2days_2050fast.xlsx')
profilesR = pd.read_excel('Net_Bus_Present_Max_Qmvar_2days_2050fast.xlsx')
DSA = DFData(profilesA)
DSR = DFData(profilesR)

def create_controllers(net, DSA, DSR, profilesA, profilesR):
    for load_idx in net.load.index:
        element_index = [load_idx]
        profilesA_name = list(profilesA)[load_idx]
        profilesR_name = list(profilesR)[load_idx]
        ConstControl(net, element='load', variable='p_mw', element_index=element_index, data_source=DSA, profile_name=profilesA_name)
        ConstControl(net, element='load', variable='q_mvar', element_index=element_index, data_source=DSR, profile_name=profilesR_name)
        
def create_output_writer(net, time_steps, output_dir):
    ow = OutputWriter(net, time_steps, output_path=output_dir, output_file_type='.xlsx', log_variables=list())
    #these variables are saved to the harddisk after / during the time series loop
    ow.log_variable('res_load', 'p_mw')
    ow.log_variable('res_bus', 'vm_pu')
    ow.log_variable('res_line', 'loading_percent')
    ow.log_variable('res_line','p_from_mw')
    ow.log_variable('res_line','p_to_mw')
    ow.log_variable('res_line', 'i_ka')
    ow.log_variable('res_trafo', 'loading_percent')
    return ow

def timeseries(output_dir, DSA, DSR, profilesA, profilesR):
    n_timesteps = 48
    create_controllers(net, DSA, DSR, profilesA, profilesR)
    time_steps = range(0, n_timesteps)
    ow = create_output_writer(net, time_steps, output_dir)
    pp.timeseries.run_timeseries(net, time_steps)
    
    

output_dir = '/Users/andrewaashish/Desktop/Thesis/ResultsMax2050Fast'
if not os.path.exists(output_dir):
    os.mkdir(output_dir)
timeseries(output_dir, DSA, DSR, profilesA, profilesR)
  
pp.diagnostic(net, report_style='detailed', warnings_only=False, return_result_dict=True, overload_scaling_factor=0.001, min_r_ohm=0.001, min_x_ohm=0.001, min_r_pu=1e-05, min_x_pu=1e-05, nom_voltage_tolerance=0.3, numba_tolerance=1e-05)
# voltage results
import os
import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
vm_pu_file = os.path.join(output_dir, "res_bus", "vm_pu.xlsx")
vm_pu = pd.read_excel(vm_pu_file, index_col=0)

# Specify the column indices you want to plot (58, 44, and 45)
column_indices_to_plot = [25, 58, 26, 29]  # Add the desired column indices

# Plot the selected columns
for column_index in column_indices_to_plot:
    vm_pu.iloc[:, column_index].plot(label=str(column_index))

# Customize the plot
plt.xlabel("time step")
plt.ylabel("voltage mag. [p.u.]")
plt.title("Voltage Magnitude")
plt.grid()

# Position the legend horizontally below the x-axis
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=len(column_indices_to_plot))

plt.show()

# line loading results
ll_file = os.path.join(output_dir, "res_line", "loading_percent.xlsx")
line_loading = pd.read_excel(ll_file, index_col=0)
line_loading.plot(label="line_loading", legend = '')
plt.xlabel("time step")
plt.ylabel("line loading [%]")
plt.title("Line Loading")
plt.grid()
plt.show()

# line losses results
pfrom_file = os.path.join(output_dir, "res_line", "p_from_mw.xlsx")
pto_file = os.path.join(output_dir, "res_line", "p_to_mw.xlsx")
powerinput = pd.read_excel(pfrom_file, index_col=0)
poweroutput = pd.read_excel(pto_file, index_col=0)
losses = (powerinput + poweroutput)*1000
losses.plot(label="losses", legend = '')
plt.xlabel("time step")
plt.ylabel("Losses [kW]")
plt.title("Losses")
plt.grid()
plt.show()

# trafo loading results
tl_file = os.path.join(output_dir, "res_trafo", "loading_percent.xlsx")
trafo_loading = pd.read_excel(tl_file, index_col=0)

# Specify the column indices you want to plot
column_indices_to_plot = [4, 7, 22, 20,]  # Add the desired column indices

# Plot the selected columns
for column_index in column_indices_to_plot:
    trafo_loading.iloc[:, column_index].plot(label=str(column_index))

# Customize the plot
plt.xlabel("time step")
plt.ylabel("trafo loading [%]")
plt.title("Transformer Loading")
plt.grid()

# Position the legend horizontally below the x-axis
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=len(column_indices_to_plot))

plt.show()

# load results
load_file = os.path.join(output_dir, "res_load", "p_mw.xlsx")
load = pd.read_excel(load_file, index_col=0)
load.plot(label="load", legend = '')
plt.xlabel("time step")
plt.ylabel("P [MW]")
plt.grid()
plt.show()

print(net.bus_geodata)
print(net.bus)
print(net.line)
print(net.load)
pp.runpp(net)
print(net.res_bus)
print(net.res_line)
print(net.res_trafo)
print(net)