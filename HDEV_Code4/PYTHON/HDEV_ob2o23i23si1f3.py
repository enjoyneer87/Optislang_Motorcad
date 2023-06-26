# %% [markdown]
# # [Packages] Import

# %%
import win32com.client
import os
import numpy as np
import time
from os import getcwd
from os.path import join, dirname, exists
from math import pi, sqrt
from scipy.io import loadmat
import re

# %% [markdown]
# # [Functions] Define (User)
# 
# 
# 

# %% [markdown]
# ## Calc Dimension

# %% [markdown]
# <a id='calcImplicitToothWidthRatio'></a>
# ### calcImplicitToothWidthRatio

# %%
def calcImplicitToothWidthRatio(BackIronThickness,i_YtoT,p_Tooth_Tip_Depth,i_Stator_OD,i_Split_Ratio,p_Stator_Slots):
    # stator Dimension
    Rint = (i_Stator_OD *i_Split_Ratio)/2    
    import math
    slot_pitch=360/p_Stator_Slots  # [in angle]
    slot_pitch_rad  =  math.radians(slot_pitch)
    ## Tooth Width Calc
    ToothWidth=BackIronThickness/i_YtoT
    # Radius
    Radius_ToothWidth=Rint+p_Tooth_Tip_Depth-0.05 
    # Angle 
    Angle_Radian_ToothWidth=ToothWidth/Radius_ToothWidth
    imp_Tooth_Width_Ratio=Angle_Radian_ToothWidth/slot_pitch_rad

    return Rint,slot_pitch,ToothWidth,Angle_Radian_ToothWidth,imp_Tooth_Width_Ratio

# %% [markdown]
# <a id='calcBackIronThickness'></a>
# ### calcBackIronThickness

# %%
def calcBackIronThickness(i_Stator_OD,i_Split_Ratio,i_Depth_Slot_Ratio,p_MinThicknessBackIron):
    Rext=i_Stator_OD/2 
    Rint = (i_Stator_OD *i_Split_Ratio)/2
    ThicknessStatorRegion=Rext-Rint
    BackIronThickness=(1-i_Depth_Slot_Ratio)*(ThicknessStatorRegion-p_MinThicknessBackIron)
    return BackIronThickness

# %% [markdown]
# <a id='calcBackIronThickness'></a>
# ### CalcSlotOpening 

# %%
def plot_arc(radius, start_angle, end_angle):
    import numpy as np
    theta = np.linspace(start_angle, end_angle, 100)
    x = radius * np.cos(theta)
    y = radius * np.sin(theta)
    # plt.plot(x, y)
    # plt.axis('equal')

def calculate_arc_length(radius, angle):
    return radius * angle

def calculate_distance(radius, angle):
    import numpy as np
    return 2 * radius * np.sin(angle / 2)

def CalcSlotOpening(Rint,p_Tooth_Tip_Depth,slot_pitch,Angle_Radian_ToothWidth):
    import numpy as np
    # Example usage
    radius = Rint+p_Tooth_Tip_Depth
    start_angle = 0
    # end_angle = 2.5* np.pi / 180
    end_angle=(slot_pitch*np.pi/180-Angle_Radian_ToothWidth)
    plot_arc(radius, start_angle, end_angle)

    arc_length = calculate_arc_length(radius, end_angle - start_angle)
    distance = calculate_distance(radius, end_angle - start_angle)
    # # Calculate the coordinates of the two points on the arc
    # x_start = radius * np.cos(start_angle)
    # y_start = radius * np.sin(start_angle)
    # x_end = radius * np.cos(end_angle)
    # y_end = radius * np.sin(end_angle)
    # # Plot a straight line representing the distance
    # plt.plot([x_start, x_end], [y_start, y_end], 'r--')
    # plt.show()
    return distance



# %% [markdown]
# <a id='fun_YtoT'></a>
# ### [Not In Use] fun_YtoT

# %%

### Geometry

# def fun_YtoT(i_YtoT, i_Tooth_Width):
#     res = i_Tooth_Width*i_YtoT
#     return res


# %% [markdown]
# ## Calc input

# %% [markdown]
# <a id='fun_MaxTorque'></a>
# ### fun_MaxTorque

# %%
def fun_MaxTorque(motdata,p_Speed_Peak_Array,i_lineCurrentLabpk,i_TurnLab):
    mot_file_dir=motdata[0]
    Mat_File_Name=motdata[1]
    
    mcApp.CalculateMagnetic_Lab()                                 # Run calculation
    mcApp.SetVariable("LabThermalCoupling", 0)                    # Coupling with Thermal
    Mat_File_Path     = join(mot_file_dir, 'Lab', Mat_File_Name)  # Point to the *.mat file
    Mat_File_Data     = loadmat(Mat_File_Path)                    # Load data from the *.mat file
    Mat_File_Speed    = Mat_File_Data['Speed']                    # Load speed data
    Mat_File_Torque   = Mat_File_Data['Shaft_Torque']             # Load shaft torque data    
    Mat_File_Power    = Mat_File_Data['Shaft_Power']              # Load shaft power data
    Mat_File_Torque   = Mat_File_Torque.flatten()                 # Necessary to be read by list_2_list_variant()
    Mat_File_Power    = Mat_File_Power.flatten()  
    Mat_File_Speed    = Mat_File_Speed.flatten()
  # Extract specific values
    Peak_Power_Array = np.zeros(len(p_Speed_Peak_Array))   # Peak power array initialisation
    Peak_Torque_Array = np.zeros(len(p_Speed_Peak_Array))  # Peak power array initialisation
    for i in range(len(p_Speed_Peak_Array)):
        Ind_Speed = (np.abs(Mat_File_Speed - p_Speed_Peak_Array[i])).argmin()   # Index corresponding to the required speed
        Peak_Power_Array[i]  = Mat_File_Power[Ind_Speed]                        # Peak power at the required speed 
        Peak_Torque_Array[i] = Mat_File_Torque[Ind_Speed]                       # Peak torque at the required speed 
    o_Peak_Torque_1700rpm    = Peak_Torque_Array[0]      # In [Nm] 
    # o_Peak_Power_4krpm      = Peak_Power_Array[1]       # In [kW]    
    return o_Peak_Torque_1700rpm


# %% [markdown]
# <a id='fun_Find_Ipk_4Trq65C_mk_dat'></a>
# ### fun_Find_Ipk_4Trq65C_mk_dat

# %%
def fun_Find_Ipk_4Trq65C_mk_dat(ext_Duty_Cycle,i_TurnLab):
    mcApp.SetVariable('TurnsCalc_MotorLAB', i_TurnLab)                 # Turns per coil
    mcApp.SetVariable("LabThermalCoupling", 0)                         # Coupling with Thermal
    mcApp.SetVariable("OpPointSpec_MotorLAB", 0)                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("SpeedDemand_MotorLAB", ext_Duty_Cycle[1])                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("TorqueDemand_MotorLAB", ext_Duty_Cycle[2])                       # 0- Torque 4-Max temperature definition  
    mcApp.CalculateOperatingPoint_Lab()                             # Operating point calculation
    ex, LabOpPoint_ShaftTorque = mcApp.GetVariable("LabOpPoint_ShaftTorque")    # Get shaft torque value
    ex, ipk = mcApp.GetVariable("LabOpPoint_StatorCurrent_Line_Peak")    # Get ipk
    ex, beta = mcApp.GetVariable("LabOpPoint_PhaseAdvance")    # Get beta
    before,after =fun_Save_Duty_Cycle_Change_I(ext_Duty_Cycle,ipk)     

    # Get peak Current check value
    mcApp.ShowMagneticContext()          
    mcApp.SetVariable('PeakCurrent',ipk)
    ex, ipk_check = mcApp.GetVariable("PeakCurrent")    
    #  # Get rms current density of OP1 value
    # if ipk_check == ipk:
    #     ex, o_current_densityRMS = mcApp.GetVariable("RMSCurrentDensity")
    o_current_densityRMS,o_copper_area = fun_current_density_cal(ipk_check,i_TurnLab)
    
    return ipk, beta, LabOpPoint_ShaftTorque, after, o_current_densityRMS,o_copper_area


# %% [markdown]
# <a id='fun_Save_Duty_Cycle_Change_I'></a>
# ### fun_Save_Duty_Cycle_Change_I
# 

# %%
def fun_Save_Duty_Cycle_Change_I(ext_Duty_Cycle,Ipk):
    ## Load reference Duty Cycle with ext_Duty_Cycle_name
    ref_Duty_Cycle=join(dirname(dirname(OSL_PROJECT_DIR)), 'DutyCycleData',ext_Duty_Cycle[0])+'.dat'
    mcApp.loadDutyCycle(ref_Duty_Cycle)

    ## change current
    ex, Duty_Cycle_Num=mcApp.GetVariable('Duty_Cycle_Num_Periods')
    ex, Duty_Cycle_Current_before=mcApp.GetVariable('Duty_Cycle_Current_Start')
    for i in range(0,Duty_Cycle_Num):   
        mcApp.SetArrayVariable('Duty_Cycle_Current_Start',i,Ipk/sqrt(2))
    ex, Duty_Cycle_Current_After=mcApp.GetVariable('Duty_Cycle_Current_Start')
    
    ## Save Duty Cycle current
    if not exists(join(mot_file_dir,'DutyCycleData')):
        os.mkdir(join(mot_file_dir,'DutyCycleData'))    
    ext_Duty_Cycle_new=join(mot_file_dir,'DutyCycleData',ext_Duty_Cycle[0])+'_new.dat'
    mcApp.SaveDutyCycle(ext_Duty_Cycle_new)
    return Duty_Cycle_Current_before,Duty_Cycle_Current_After


# %% [markdown]
# <a id='fun_Calc_Temp_ext_duty_rename'></a>
# ### fun_Calc_Temp_ext_duty_rename

# %%

# Calculate Temp Rise with duty cycle
def fun_Calc_Temp_ext_duty_rename(ext_Duty_Cycle,i_TurnLab,Thermal_CouplingType):
    ext_Duty_Cycle_new=join(mot_file_dir,'DutyCycleData',ext_Duty_Cycle[0])+'_new.dat'
    mcApp.LoadDutyCycle(ext_Duty_Cycle_new)
    mcApp.SetVariable('TurnsCalc_MotorLAB', i_TurnLab)            # Turns per coil
    mcApp.SetVariable("LabThermalCoupling_DutyCycle", Thermal_CouplingType)                         # Coupling with Thermal
    mcApp.SetVariable("LabThermalCoupling", Thermal_CouplingType)                         # Coupling with Thermal
    mcApp.SetVariable('InitialTransientTemperatureOption',3)
    mcApp.CalculateDutyCycle_Lab()
    fun_Rename_Matfile_Lab_Duty(ext_Duty_Cycle)      


# %% [markdown]
# <a id='fun_Rename_Matfile_Lab_Duty'></a>
# ### fun_Rename_Matfile_Lab_Duty
# 

# %%
   
def fun_Rename_Matfile_Lab_Duty(ext_Duty_Cycle):
    ex, motpath=mcApp.GetVariable("CurrentMotFilePath_MotorLAB")
    motpath=re.sub(".mot","",motpath)
    Lab_path=motpath+'/Lab/'
    os.chdir(Lab_path)
    rename_matfile=ext_Duty_Cycle[0]+'_lab_result.mat'
    if os.path.exists('MotorLAB_drivecycledata.mat'):
        os.rename('MotorLAB_drivecycledata.mat',rename_matfile)
      


# %% [markdown]
# <a id='fun_load_matfile_by_OP'></a>
# ### fun_load_matfile_by_OP

# %%
   
# Check Temp Rise with duty cycle
def fun_load_matfile_by_OP(ext_Duty_Cycle):
        ## This mat file consisted of Ndarray when we are using loadmat function
    ex, motpath=mcApp.GetVariable("CurrentMotFilePath_MotorLAB")
    motpath=re.sub(".mot","",motpath)
    Lab_path=motpath+'/Lab/'
    os.chdir(Lab_path)
    name_matfile=ext_Duty_Cycle[0]+'_lab_result.mat'
    Mat_File_Data=loadmat(name_matfile)
    return Mat_File_Data
## last function for Duty Cycle check


# %% [markdown]
# <a id='fun_Check_Temp_Rise_allComponent'></a>
# ### fun_Check_Temp_Rise_allComponent

# %%
def fun_Check_Temp_Rise_allComponent(ext_Duty_Cycle):
    ## init
    init_final_temp=[]
    # ## change name of mat file
    # fun_rename_matfile_lab_duty(ext_Duty_Cycle)
    ## load mat file
    list_from_mat=fun_load_matfile_by_OP(ext_Duty_Cycle)
    
    ## 
    mat_temp=[list_from_mat.get(key) for key in list_from_mat.keys() if 'Temp' in key]
    mat_temp_key=[key for key in list_from_mat.keys() if 'Temp' in key]

    ## check temp rise
    for i in range(len(mat_temp)):
        temp=mat_temp[i].ravel().tolist()
        init_final_temp.append((temp[0],temp[-1]))
        check_temp=max(max(init_final_temp))
    
    ## 
    max_temp_key=[key for key in mat_temp_key if max(list_from_mat.get(key))==check_temp]
    dic_init_final_temp=dict(zip(mat_temp_key,init_final_temp))
    
    return check_temp,max_temp_key, dic_init_final_temp


# %% [markdown]
# <a id='fun_OP_temp_contraints'></a>
# ### fun_OP_temp_contraints

# %%

# set value for contrainsts
def fun_OP_temp_contraints(ext_Duty_Cycle,i_TurnLab,Thermal_CouplingType):
    Op_i,OP_beta,Op_LabOpPoint_ShaftTorque,Op_after,current_density,o_copper_area =fun_Find_Ipk_4Trq65C_mk_dat(ext_Duty_Cycle,i_TurnLab)
    fun_Calc_Temp_ext_duty_rename(ext_Duty_Cycle,i_TurnLab,Thermal_CouplingType)
    o_OP_max_temp,max_pos,temp_dic=fun_Check_Temp_Rise_allComponent(ext_Duty_Cycle)
    return o_OP_max_temp,max_pos,temp_dic, Op_i,OP_beta,Op_after,Op_LabOpPoint_ShaftTorque,current_density,o_copper_area


# %% [markdown]
# <a id='fun_Driving_Duty'></a>
# ### fun_Driving_Duty Set Initial Temp

# %%

def fun_Driving_Duty(ext_Duty_Cycle,turns):
    ref_Duty_Cycle=join(dirname(dirname(OSL_PROJECT_DIR)), 'DutyCycleData',ext_Duty_Cycle[0])+'.dat'
    mcApp.LoadDutyCycle(ref_Duty_Cycle)
    mcApp.SetVariable('TurnsCalc_MotorLAB', turns)            # Turns per coil
    mcApp.SetVariable("LabThermalCoupling_DutyCycle", 0)                         # Coupling with Thermal Duty Cycle
    mcApp.SetVariable("LabThermalCoupling", 0)                         # Coupling with Thermal
    mcApp.SetVariable('InitialTransientTemperatureOption',4)
    mcApp.SetVariable('InitialHousingTemperature',65)
    mcApp.SetVariable('InitialHousingTemperature',65)
    mcApp.SetVariable('InitialStatorTemperature',140)
    mcApp.SetVariable('InitialWindingTemperature',160)
    mcApp.SetVariable('InitialRotorTemperature',100)
    mcApp.SetVariable('InitialMagnetTemperature',130)
    mcApp.CalculateDutyCycle_Lab()
    # Calculation & Post processing
    #ex, o_WLTP3_Eff = mcApp.GetVariable("DutyCycleAverageEfficiency_EnergyUse")   # Get efficiency value 
    ex, o_Wh_Loss = mcApp.GetVariable("DutyCycleTotalLoss")   # Get efficiency value 
    ex, o_Wh_Shaft = mcApp.GetVariable("DutyCycleTotalEnergy_Shaft_Output")   # Get efficiency value 
    ex, o_Wh_input = mcApp.GetVariable("DutyCycleTotalEnergy_Electrical_Input")   # Get efficiency value 
    # o_max_temp,max_pos,temp_dic=fun_Check_Temp_Rise_allComponent(ext_Duty_Cycle)
    return o_Wh_Loss,o_Wh_Shaft,o_Wh_input


# %% [markdown]
# ## Calc output

# %% [markdown]
# <a id='fun_output_define'></a>
# ### fun_output_define
# 

# %%

## 
def fun_output_define():
    o_LabCurrentJ=0
    # OP point check
    o_Op1_Jrms =0
    o_Op2_Jrms =0
    o_Op3_Jrms =0
    o_Op1_ipk               =0
    o_Op2_ipk               =0
    o_Op3_ipk               =0
    o_Op1_ShaftToruqe       =0
    o_Op2_ShaftToruqe       =0
    o_Op3_ShaftToruqe       =0
    # Temp rise test (Duty Cycle)
    o_Op2_max_temp          =0
    # o_OP3_max_temp          =0
    # Driving Duty Cycle
    o_Wh_Loss           =0
    o_Wh_Shaft          =0
    o_Wh_input          =0
    # Torque density 
    o_Maxtorque=0
    o_TorqueVolumeDensity =0
    o_TorqueWeightDensity =0
    o_Op2_copper_area   =0
    o_Op1_copper_area   =0
    o_Op3_copper_area   =0
    # Weight
    o_Weight_Act         = 0.
    o_Weight_Mag         = 0.
    o_Weight_Rot_Core    = 0.
    o_Weight_Stat_Core   = 0.
    o_Weight_Wdg         = 0.
    
    outputList=['o_LabCurrentJ',o_LabCurrentJ]
    outputList.extend(['o_Op1_Jrms',o_Op1_Jrms])
    outputList.extend(['o_Op2_Jrms',o_Op2_Jrms])
    outputList.extend(['o_Op3_Jrms',o_Op3_Jrms])
    outputList.extend(['o_Op1_ipk',o_Op1_ipk])
    outputList.extend(['o_Op2_ipk',o_Op2_ipk])
    outputList.extend(['o_Op3_ipk',o_Op3_ipk])
    outputList.extend(['o_Op1_ShaftToruqe',o_Op1_ShaftToruqe])
    outputList.extend(['o_Op2_ShaftToruqe',o_Op2_ShaftToruqe])
    outputList.extend(['o_Op3_ShaftToruqe',o_Op3_ShaftToruqe])
    outputList.extend(['o_Op2_max_temp',o_Op2_max_temp])
    # outputList.extend(['o_OP3_max_temp',o_OP3_max_temp])
    outputList.extend(['o_Wh_Loss',o_Wh_Loss])
    outputList.extend(['o_Wh_Shaft',o_Wh_Shaft])
    outputList.extend(['o_Wh_input',o_Wh_input])
    outputList.extend(['o_Maxtorque',o_Maxtorque])
    outputList.extend(['o_TorqueVolumeDensity',o_TorqueVolumeDensity])
    outputList.extend(['o_TorqueWeightDensity',o_TorqueWeightDensity])
    outputList.extend(['o_Op2_copper_area',o_Op2_copper_area])
    outputList.extend(['o_Op1_copper_area',o_Op1_copper_area])
    outputList.extend(['o_Op3_copper_area',o_Op3_copper_area])
    outputList.extend(['o_Weight_Act',o_Weight_Act])
    outputList.extend(['o_Weight_Mag',o_Weight_Mag])
    outputList.extend(['o_Weight_Rot_Core',o_Weight_Rot_Core])
    outputList.extend(['o_Weight_Stat_Core',o_Weight_Stat_Core])
    outputList.extend(['o_Weight_Wdg',o_Weight_Wdg])
    return outputList
        #return o_Maxtorque,o_TorqueVolumeDensity,o_TorqueWeightDensity,o_OP3_max_temp,o_LabCurrentJ,o_Op1_Jrms,o_Op2_Jrms,o_Op3_Jrms,o_Op1_ipk,o_Op2_ipk,o_Op3_ipk,o_Op1_ShaftToruqe,o_Op2_ShaftToruqe,o_Op3_ShaftToruqe,o_Wh_Loss,o_Wh_Shaft,o_Wh_input,o_Op2_max_temp,o_Torque_Density,o_Weight_Act,o_Weight_Mag,o_Weight_Rot_Core,o_Weight_Stat_Core,o_Weight_Wdg,o_Active_Volume,o_Op1_copper_area,o_Op2_copper_area,o_Op3_copper_area


# %% [markdown]
# <a id='fun_current_density_cal'></a>
# ### fun_current_density_cal

# %%

## Simple Calculation
def fun_current_density_cal(ipk_check,turns):
    ex, o_copper_area = mcApp.GetVariable('Copper_Area')    
    ex,p_Parallel_Path=mcApp.GetVariable('ParallelPaths')
    o_current_densityRMS = (ipk_check/sqrt(2))* turns / p_Parallel_Path /o_copper_area
    return  o_current_densityRMS,o_copper_area


# %% [markdown]
# <a id='fun_Turn_byAmpT'></a>
# ### [Not Used]fun_Turn_byAmpT

# %%

# def fun_Turn_byAmpT(i_AmpT,i_BuildLineCurrentRMS):
#     res = i_AmpT/i_BuildLineCurrentRMS
#     return res


# %% [markdown]
# <a id='fun_Ipk_beta_by_Trq'></a>
# ### fun_Ipk_beta_by_Trq
# 

# %%

def fun_Ipk_beta_by_Trq():
    ex, ipk = mcApp.GetVariable("LabOpPoint_StatorCurrent_Line_Peak")    # Get shaft torque value
    ex, beta = mcApp.GetVariable("LabOpPoint_PhaseAdvance")    # Get shaft torque value
    return ipk, beta


# %% [markdown]
# <a id='fun_scail_J_by_Turn'></a>
# ### fun_scail_J_by_Turn
# 

# %%

def fun_scail_J_by_Turn(op_current_densityRMS,p_init_Turns_Coil,turns):
    # o_op_J = op_current_density*turns/p_init_Turns_Coil #false
    o_op_J = op_current_densityRMS
    return o_op_J


# %% [markdown]
# <a id='fun_Machine_Length'></a>
# ### fun_Machine_Length

# %%

# Machine length
def fun_Machine_Length(p_EndSpace_Height,p_Wdg_Overhang_F, p_Wdg_Overhang_R, i_Active_Length):
    res = i_Active_Length + 2*(p_EndSpace_Height) + p_Wdg_Overhang_F+p_Wdg_Overhang_R
    return res


# %% [markdown]
# <a id='fun_Active_Volume'></a>
# ### fun_Active_Volume

# %%
## Active volume
def fun_Active_Volume(Stator_OD, Active_Length):  
    res = pi*Stator_OD**2/4*Active_Length*1e-9
    return res


# %% [markdown]
# <a id='fun_TorqueVolumeDensity'></a>
# ### fun_TorqueVolumeDensity

# %%

### Performances
# Torque density
def fun_TorqueVolumeDensity(Torque, Volume):           
    res = Torque/Volume*1e-3
    return res


# %% [markdown]
# <a id='fun_TorqueWeightDensity'></a>
# ### fun_TorqueWeightDensity

# %%
def fun_TorqueWeightDensity(Torque, weight):           
    res = Torque/weight*1e-3
    return res


# %% [markdown]
# <a id='fun_Stress_Safety'></a>
# ### fun_Stress_Safety

# %%

 # Stress safety factor
def fun_Stress_Safety(Rotor_Yield, Stress_Max):            
    res = Rotor_Yield/Stress_Max
    return res


# %% [markdown]
# <a id='funCalcEMFFundLab'></a>
# ### funCalcEMFFundLab

# %%
def funCalcEMFFundLab(SpeedBEMF,scaleCoeff):        
    ex, NoloadPsiD=mcApp.GetArrayVariable ('PsiDModel_Lab',0) # [Vs] Unit  
    ex, NoloadPsiQ=mcApp.GetArrayVariable ('PsiQModel_Lab',0) # [Vs] Unit  

    ex, p=mcApp.GetVariable('Pole_Number')
    Wr_test = 2 * pi * SpeedBEMF / 60 *p / 2;  
    emfD=Wr_test*NoloadPsiD
    emfQ=Wr_test*NoloadPsiQ
    phaseEmfFund=sqrt(emfD**2+emfQ**2)
    LineEmfFund=phaseEmfFund*sqrt(3)
    LineEmfFund=scaleCoeff*LineEmfFund
    return LineEmfFund


# %% [markdown]
# # [Start Code] Running Mode & Working directories

# %% [markdown]
# ## Mode detection

# %%
if 'OSL_PROJECT_DIR' in locals():                       # Working in optiSLang 
    within_OSL = True
else:
    within_OSL = False                                  # Working in IDE console 
    OSL_REGULAR_EXECUTION = False
    OSL_DESIGN_NAME = 'Design0001'   
    OSL_DESIGN_NO   = 1.0    

if within_OSL:                                          # Working in optiSLang
    if OSL_REGULAR_EXECUTION:
        run_mode = 'OSL_run'                            # OSL-runtime mode
    else:
        run_mode = 'OSL_setup'                          # OSL-setup mode
    from pyvariant import list_list_2_variant_signal
    
else:                                                   # Working in IDE
    run_mode = 'IDE_run'                                # IDE-run mode

# Directories
if run_mode.startswith('OSL'):                                 # Working in optiSLang 
    print('[INFO] Running in OSL environment')
    wdir = OSL_DESIGN_DIR
    refdir = join(dirname(dirname(OSL_PROJECT_DIR)), 'MCAD')
    print('[INFO] Working directory: ', wdir)
    print('[INFO] Reference MCAD directory: ', refdir)
else:                                                          # Working in IDE console
    print('[INFO] Running in IDE console')
    refdir = join(dirname(dirname(__file__)), 'MCAD')
    wdir = join(getcwd(), 'test_run')
    if not exists(wdir):
        os.mkdir(wdir)
    print('[INFO] Working directory: ', wdir)
    print('[INFO] Reference MCAD directory: ', refdir)



# %% [markdown]
# # Input

# %% [markdown]
# ## Common Input (You can Modifiy the Cells with User written)

# %%
Visible_Opt     = 1.           # Set Motor-CAD visible
Message_Display = 2.           # Display all pop-up messages 
Save_Prompt     = 1.           # Never prompt to save file

# %% [markdown]
# ### Motor-CAD options (User)

# %%
Design_Name     = "HDEV_Model2"   # Reference Motor-CAD design
### Geometry
p_Pole_Pair            = 6.    # Number of rotor poles
p_Stator_Slots         = 72.   # Number of stator slots


# %% [markdown]
# #### [TBC] Fixed(TBC PreDefined) Variable (Read initial Value From MOT File)

# %%
### Stator
# ### Absoulute Input (left table)
i_Slot_Corner_Radius    = 1.4     
#i_Tooth_Width  = 6     
p_Tooth_Tip_Depth       = 1    # Tooth tip depth 
p_Tooth_Tip_Angle       = 20   # Tooth tip angle 
#### Ratio(Stator w Hierachy)
p_MinThicknessBackIron = 0.1
# p_Slot_Depth_Ratio      =   1   

# ### Rotor 
# #### Absoulute Input (left table)
#     # Notch_Depth=0 
# Magnet_Layers=2 
# #### Ratio  (Rotor w Hierachy)
p_Airgap_Mecha         = 1.    # Mechanical airgap
# p_Mag_Clear            = 0   # Magnet clearance

### other
p_EndSpace_Height      = 24.5   # Space between winding ends and caps
p_Wdg_Overhang_F         = 56.   # Winding overhang height
p_Wdg_Overhang_R        = 65.   # Winding overhang height

### Winding
p_Coils_Slot    = 1.    # Number of coils going through each slot
p_Parallel_Path = 4.    # Number of parallel paths per phase
p_Slot_Fill     = 0.5325   # Copper slot fill factor
p_init_Turns_Coil =11.   # Initial number of turns per coil

# ### Materials
p_Yield_Rotor  = 460.   # Rotor core yield strength
p_Temp_Wdg_Max = 180.   # Maximum winding temperature
p_Temp_Mag_Max = 140.   # Maximum magnet temperature

### Performance
p_Speed_Max        = 6000.   # Maximum operating speed


# %% [markdown]
# ### Calculation Option (User)

# %%
p_Speed_Lab_Step    = 100.                              # Speed step used in Lab
p_Speed_Peak_Array  = np.array([1700.])    # Speeds for peak performance calculation 
p_Speed_Cont_Array  = np.array([1700., 4000.])          # Speeds for continuous performance calculation
p_Torque_Pts        = 90                                # Timesteps per cycle for torque calculation   
                                           
### Post-processing
Pic_Export = 1      # Export geometry snapshots (0: No  1: Yes)

### Dependent parameters
Speed_Max_Rad = pi*p_Speed_Max/30                                               # Maximum speed in radians
Speed_Lab     = np.arange(0, p_Speed_Max + p_Speed_Lab_Step, p_Speed_Lab_Step)  # Speed vector in Lab
Speed_Lab     = Speed_Lab.tolist()                                              # Required for signal generation
Speed_Lab_Len = len(Speed_Lab)                                                  # Required for signal generation

#calc option
OP2thermalCouplingType=2 # 0: no thermal coupling 1: Loss-> Thermal # closed coupled
OP3thermalCouplingType=1 # 0: no thermal coupling 1: Loss-> Thermal # closed coupled

ModelType_MotorLAB     =2             # Saturation model type: 0 Fixed Inductance 1: singel step 2: Full Cycle
SatModelPoints_MotorLAB=1           # Saturation model: 0 - coarse 1- fine resolution (30 points)
LossModel_Lab          =1        # Loss model type: 1-FEA 2 -custom
BuildSatModel_MotorLAB=1          # Activate saturation model               
BuildLossModel_MotorLAB=1         ## Activate loss model                
CalcTypeCuLoss_MotorLAB=3         # 0 DC only 1 DC+AC(User) 2 DC+AC (FEA single Point) 3 DC+AC (FEA Map)      
ACLossGeneratorMethod_Lab =1      
ProximityLossModel = ACLossGeneratorMethod_Lab   
IronLossCalc_Lab=3               # 0Neglect 1 OC+SC(User) 2 OC+SC (FEA single Point) 3 (FEA Map)
LabModel_MagnetLoss_Method=3   #0 Neglect 1 User Defined 2 OC+SC (FEA single Point) 3 (FEA Map)         
MagnetLossCalc_Lab =3          # 0 Neglect 1 User Defined 2 OC+SC (FEA single Point) 3 (FEA Map)    


# %% [markdown]
# ## [Setup Mode]Input parameters for testing in IDE or initialisation in OSL Python node
#      (just Setup variable which will be changed by Optislang while OSL_run)
# 

# %%
if run_mode in ['OSL_setup', 'IDE_run']:
  
    i_Active_Length     = 130.   # Active length
    
    ### Performance
    i_BuildLineCurrentRMS =  742.4621 # 1050   # Maximum RMS line current  900Apk 636.3961030678927A rms but 1050Apk 742.4621A rms
    # i_AmpT_rms=                 # Maximum Ampere turn current 2.75T*900ApkA  1750 rms 1050Apk 2.75T 742.4621A rms
    # p_AmpT_rms= 1555.6
    
    ### Winding
    i_lineCurrentLabpk = 0
    p_init_Turns_Coil =11          # for model build use fixed values for initialisation change by i_TurnLab
    


# %% [markdown]
# ### Stator 
# 

# %%
    # i_Tooth_Width           = 6
    i_Stator_OD             = 400  # Stator outer diameter
    ## Lab settings
    i_TurnLab = 11
    #### Ratio(Stator w Hierachy)
    # Bore
    i_Split_Ratio                       =0.705 
    # Depth  
    i_Depth_Slot_Ratio  =0.7
    ##exlicit [Not In MCAD variable But Optislang]
    BackIronThickness=calcBackIronThickness(i_Stator_OD,i_Split_Ratio,i_Depth_Slot_Ratio,p_MinThicknessBackIron)
    i_YtoT                              =2.5                                           #ratio user defined YtoT 
    # Tooth Width
    ## implicit (Not in Optislang But MCAD)
    Rint,slot_pitch,ToothWidth,Angle_Radian_ToothWidth,imp_Tooth_Width_Ratio=calcImplicitToothWidthRatio(BackIronThickness,i_YtoT,p_Tooth_Tip_Depth,i_Stator_OD,i_Split_Ratio,p_Stator_Slots)
    # Slot Opening                                                                    
    i_Slot_Op_Ratio        =0.8     

    

# %% [markdown]
# ### Rotor 
# 

# %%
### Rotor 
    #### Absoulute Input (left table)
        # Notch_Depth=0 
    Magnet_Layers=2

    # # Layer 1 - insider
    L1_Magnet_Thickness     =5.2
    L1_Bridge_Thickness     =1.8
    L1_Pole_V_angle         =112
    L1_Magnet_Post          =1.5
    L1_Magnet_Separation    =6.2
    L1_Magnet_Segments      =1
    # L1 Mag Gap Inner
    # L1 Mag Gap Outer

    # # Layer 2
    L2_Magnet_Thickness      =5.4 
    L2_Bridge_Thickness      =1.5 
    L2_Magnet_Post           =0 
    L2_Magnet_Separation     =0 
    L2_Magnet_Segments       =1 

    #### Ratio  (Rotor w Hierachy)
    L1_Pole_Arc                            = 0.90           
    L1_Web_Thickness                   = 0.15           
    L1_Magnet_Bar_Width            = 0.92           
    L1_Web_Length              = 0.104          
    L2_Pole_Arc                        = 0.25           
    L2_Web_Thickness               = 0.69           
    L2_Magnet_Bar_Width        = 0.8             
    L2_Web_Length          = 0.064       

    ### Duty Cycle Study
    #i_Gear_Ratio = 7.


# %% [markdown]
# # Real run Setting: (start If) 'IDE_run' mode or 'OSL_run' mode
# 

# %%
if run_mode.endswith('run'):   


# %% [markdown]
# ## [precal with function] Geometry parameters
# [Go to cell fun_Machine_Length](#fun_Machine_Length)
# 

# %%
    Machine_Length = fun_Machine_Length(p_EndSpace_Height, p_Wdg_Overhang_R,p_Wdg_Overhang_R, i_Active_Length)
    # Air_Pocket     = fun_Air_Pocket(i_Mag_Thick, p_Mag_Clear)


# %% [markdown]
# ## Load [reference] Motor-CAD file
# 

# %%
    mcApp = win32com.client.Dispatch("MotorCAD.AppAutomation")  # Launch Motor-CAD application
    # mcApp=MotorCAD()
    mcApp.Visible = Visible_Opt                                 # Set Motor-CAD visible or not
    mcApp.SetVariable('MessageDisplayState', Message_Display)   # Set state of message display 
    mcApp.SetVariable("SavePrompt", Save_Prompt)                # Remove autosave function or not
    myPID = os.getpid()											# Pass this process id to Motor-CAD so that Motor-CAD will 
    print(myPID)                                                # close when this process completes
    mcApp.SetVariable("OwnerProcessID", myPID)
    mcApp.ShowMessage(myPID)      
    mot_file_ref_path = join(refdir, Design_Name + ".mot")                       # Path to the reference *.mot file
    mot_file_new_path = join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + ".mot") # Path to the new *.mot file
    print("[INFO] Load reference MCAD file from: ", mot_file_ref_path)                                    
    mcApp.LoadFromFile(mot_file_ref_path)                                        # Load reference Motor-CAD file
    print("[INFO] Design file saved as: ", mot_file_new_path)                                    
    mcApp.SaveToFile(mot_file_new_path)                                          # Save in new location
    mot_file_dir = join(wdir, Design_Name + '_' + OSL_DESIGN_NAME)
    Mat_File_Name     = 'MotorLAB_elecdata.mat'                   # *.mat file automatically generated by Motor-CAD
    motdata=[mot_file_dir,Mat_File_Name]   
    

# %% [markdown]
# ## [Setting] SetVariable

# %% [markdown]
# ### Change to ratio mode to edit the geometry
# 

# %%
    mcApp.SetVariable('GeometryParameterisation', 1)   # Ratio mode in Motor-CAD 


# %% [markdown]
# ### Assign geometry parameters
# [Go to fun_Active_Volume](#fun_Active_Volume)
# 
# 

# %% [markdown]
# #### Stator
# 

# %%
    mcApp.SetVariable('Slot_Number', p_Stator_Slots)                           # Stator slots
    ### Stator
    mcApp.SetVariable('Housing_Dia', i_Stator_OD+40)                    # Stator lamination pack length
    #### Absoulute Input (left table)
    mcApp.SetVariable('Stator_Lam_Dia'      , i_Stator_OD)                        # Stator OD 
    mcApp.SetVariable('Slot_Corner_Radius'  , i_Slot_Corner_Radius)               # Slot_Corner_Radius      
    mcApp.SetVariable('Tooth_Tip_Depth'     , p_Tooth_Tip_Depth)                  # Tooth tip depth
    mcApp.SetVariable('Tooth_Tip_Angle'     , p_Tooth_Tip_Angle)                  # Tooth tip angle
    
    o_Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3] 
    
    #### Ratio(Stator w Hierachy)
    # Bore
    mcApp.SetVariable('Ratio_Bore', i_Split_Ratio)
    # Depth      
    mcApp.SetVariable("Ratio_SlotDepth_ParallelTooth"           , i_Depth_Slot_Ratio     ) 
    mcApp.SetVariable('MinBackIronThickness'                    , p_MinThicknessBackIron )        
        ##exlicit [Not In MCAD variable But Optislang]
    ex,Slot_DepthGetVar=mcApp.GetVariable('Slot_Depth')   
    BackIronThickness=calcBackIronThickness(i_Stator_OD,i_Split_Ratio,i_Depth_Slot_Ratio,p_MinThicknessBackIron)    
    print('Ratio_SlotDepth_ParallelTooth Slot Depth/Stator Lam Thickness(Yoke):',i_Depth_Slot_Ratio)
    print('Contraints Min BackIron(Yoke):',p_MinThicknessBackIron,',Dimension: SlotDepth',Slot_DepthGetVar)
        
    # Tooth Width     
    Rint,slot_pitch,ToothWidth,Angle_Radian_ToothWidth,imp_Tooth_Width_Ratio=calcImplicitToothWidthRatio(BackIronThickness,i_YtoT,p_Tooth_Tip_Depth,i_Stator_OD,i_Split_Ratio,p_Stator_Slots)
    mcApp.SetVariable("Ratio_ToothWidth"                        , imp_Tooth_Width_Ratio    )        #Ratio_ToothWidth" )      
    print('### Define Tooth Width ')
    print('i_YtoT:',i_YtoT)
    print('implicit Tooth_Width_Ratio=',imp_Tooth_Width_Ratio)
    print('Stator Bore As Contraints=',2*Rint,'Dimension=',ToothWidth)

    # Ratio_SlotOpening_ParallelTooth
    mcApp.SetVariable("Ratio_SlotOpening_ParallelTooth"         , i_Slot_Op_Ratio        )        #Ratio_SlotOpening_ParallelTooth" )      
    MaxSlotOpen=CalcSlotOpening(Rint,p_Tooth_Tip_Depth,slot_pitch,Angle_Radian_ToothWidth)
    SlotOpen=MaxSlotOpen*i_Slot_Op_Ratio
    # RatioOpenVsToothWidth=8/SlotOpen
    print('### Define Slot Opening')
    print('Slot Opening/Max Slot Opening',i_Slot_Op_Ratio)
    print('Contraints:',ToothWidth,', Dimension:Slot Opening',SlotOpen)
    


# %% [markdown]
# #### Rotor 

# %%
    ### Rotor 
    #### Absoulute Input (left table)
    mcApp.SetArrayVariable("MagnetThickness_Array"                , 0, L1_Magnet_Thickness)            # Layer 1 Magnet thickness
    mcApp.SetArrayVariable("MagnetThickness_Array"                , 1, L2_Magnet_Thickness)            # Layer 2 Magnet thickness
                
    mcApp.SetArrayVariable('BridgeThickness_Array'                , 0, L1_Bridge_Thickness)            # Layer 1 Bridge thickness 
    mcApp.SetArrayVariable('BridgeThickness_Array'                , 1, L2_Bridge_Thickness)            # Layer 2 Bridge thickness 
        
    mcApp.SetArrayVariable("PoleVAngle_Array"                     , 0, L1_Pole_V_angle)                # Layer 1 V-shape layer angle
    # mcApp.SetArrayVariable("PoleVAngle_Array"                     , 1, L2_Pole_V_angle)                # Layer 2 V-shape layer angle

    mcApp.SetArrayVariable("VShapeMagnetPost_Array"               , 0, L1_Magnet_Post)                 # Layer 1 Magnet post
    # mcApp.SetArrayVariable("VShapeMagnetPost_Array"               , 1, L2_Magnet_Post)                 # Layer 2 Magnet post

    mcApp.SetArrayVariable("MagnetSeparation_Array"               , 0, L1_Magnet_Separation)           # Layer 1 Magnet_Separation
    # mcApp.SetArrayVariable("MagnetSeparation_Array"               , 1, L2_Magnet_Separation)           # Layer 2 Magnet_Separation

    # mcApp.SetArrayVariable("VShapeMagnetSegments_Array"           , 0, L1_Magnet_Segments)             # Layer 1 Magnet_Segments
    # mcApp.SetArrayVariable("VShapeMagnetSegments_Array"           , 1, L2_Magnet_Segments)             # Layer 2 Magnet_Segments
    #### Ratio
    mcApp.SetArrayVariable(    'RatioArray_PoleArc'                    ,0,L1_Pole_Arc            )         # %Layer 1     RatioArray_PoleArc
    mcApp.SetArrayVariable(    'RatioArray_PoleArc'                    ,1,L2_Pole_Arc            )         # %Layer 2     RatioArray_PoleArc

    mcApp.SetArrayVariable(    'RatioArray_WebThickness'               ,0,L1_Pole_Arc            )         # %Layer 1     RatioArray_WebThickness
    mcApp.SetArrayVariable(    'RatioArray_WebThickness'               ,1,L2_Pole_Arc            )         # %Layer 2     RatioArray_WebThickness

    mcApp.SetArrayVariable(    'RatioArray_VWebBarWidth'               ,0,L1_Magnet_Bar_Width    )         # %Layer 1     RatioArray_VWebBarWidth
    mcApp.SetArrayVariable(    'RatioArray_VWebBarWidth'               ,1,L2_Magnet_Bar_Width    )         # %Layer 2     RatioArray_VWebBarWidth

    mcApp.SetArrayVariable(    'RatioArray_WebLength'                  ,0,L1_Web_Length          )         # %Layer 1     RatioArray_WebLength
    # mcApp.SetArrayVariable(    'RatioArray_WebLength'                  ,1,L2_Web_Length          )         # %Layer 2     RatioArray_WebLength            



# %% [markdown]
# #### etc

# %%
    #### etc
    mcApp.SetVariable('Pole_Number', 2*p_Pole_Pair)                            # Rotor poles
    mcApp.SetVariable('Airgap', p_Airgap_Mecha)                                # Mechanical airgap
     
    # mcApp.SetVariable('MinVMagnetAspectRatio', p_Mag_AspectRatio_Min)          # Minimum magnet aspect ratio
    # mcApp.SetVariable('MinMagnetSeparation', p_Mag_Separation_Min)             # Minimum separation between magnet poles  
    # mcApp.SetVariable('MinShaftSeparation', p_Shaft_Separation_Min)            # Minimum separation between shaft and magnets
    # mcApp.SetVariable("Ratio_ShaftD", p_Shaft_OD_Ratio)                        # Shaft OD ratio
    
    mcApp.SetVariable('Motor_Length', Machine_Length)     
    mcApp.SetVariable('Stator_Lam_Length', i_Active_Length)                    # Stator lamination pack length
    mcApp.SetVariable('Rotor_Lam_Length', i_Active_Length)                     # Rotor lamination pack length
    mcApp.SetVariable('Magnet_Length', i_Active_Length)                        # Magnet length
    mcApp.SetVariable('EWdg_Overhang_[R]', p_Wdg_Overhang_R)                     # End winding overhang (rear)
    mcApp.SetVariable('EWdg_Overhang_[F]', p_Wdg_Overhang_F)                     #                      (front)


# %% [markdown]
# #### Assign winding parameters

# %%
### Assign winding parameters
    mcApp.SetVariable('WindingLayers', p_Coils_Slot)                # Coils passing through a slot
    mcApp.SetVariable('MagTurnsConductor', p_init_Turns_Coil)       # Turns per coil
    mcApp.SetVariable('ParallelPaths', p_Parallel_Path)             # Parallel paths per phase
    mcApp.SetVariable('RequestedGrossSlotFillFactor', p_Slot_Fill)  # Slot fill factor


# %% [markdown]
# ### Assign initial calculation settings

# %% [markdown]
# #### Open Circuit (No load Calculation)

# %%
### Assign initial calculation settings
    mcApp.SetVariable("BackEMFCalculation", False)               # OC calculations deactivated
    mcApp.SetVariable("CoggingTorqueCalculation", False)         # Cogging torque calculation deactivated
    mcApp.SetVariable("ElectromagneticForcesCalc_OC", False)     # Maxwell forces calculation deactivated (OC)


# %% [markdown]
# #### On load

# %%
    mcApp.SetVariable("TorqueCalculation", False)                # Torque calculations deactivated
    mcApp.SetVariable("TorqueSpeedCalculation", False)           # Torque speed curve calculation deactivated
    mcApp.SetVariable("DemagnetizationCalc", False)              # Demagnetisation test deactivated
    mcApp.SetVariable("ElectromagneticForcesCalc_Load", False)   # Maxwell forces calculation deactivated (OL)

# %% [markdown]
# #### Parameters & Transient 

# %%
    mcApp.SetVariable("InductanceCalc", False)                   # Inductance calculation deactivated
    mcApp.SetVariable("BPMShortCircuitCalc", False)              # Transient short circuit calculation deactivated

# %% [markdown]
# #### Solver

# %%
    mcApp.SetVariable("MagneticSolver", 0)                       # Transient calculation enabled (0 Transient, 1 Multi-Static, 2 Reduced Multi-Static)
    mcApp.SetVariable("Lab_Threads_Enabled", True)               # Threading option for lab models enabled
    
    

# %% [markdown]
# ### [Check] the geometry is valid (exception raise)
# 

# %%
    success = mcApp.CheckIfGeometryIsValid(0)
    # [error case]
    if success == 0: 
        # [error case]If not valid, generate zero outputs instead of getting an error message in optiSLang
        # output initial values
         # Scalars
        o_LabCurrentJ=0
        # OP point check
        o_Op1_Jrms =0
        o_Op2_Jrms =0
        o_Op3_Jrms =0
        o_Op1_ipk               =0
        o_Op2_ipk               =0
        o_Op3_ipk               =0
        o_Op1_ShaftToruqe       =0
        o_Op2_ShaftToruqe       =0
        o_Op3_ShaftToruqe       =0
        # Temp rise test (Duty Cycle)
        o_Op2_max_temp          =0
        # o_OP3_max_temp          =0
        # Driving Duty Cycle
        o_Wh_Loss           =0
        o_Wh_Shaft          =0
        o_Wh_input          =0
        # Torque density 
        o_Maxtorque=0
        o_TorqueVolumeDensity =0
        o_TorqueWeightDensity =0
        o_Op2_copper_area   =0
        o_Op1_copper_area   =0
        o_Op3_copper_area   =0
        # Weight
        o_Weight_Act         = 0.
        o_Weight_Mag         = 0.
        o_Weight_Rot_Core    = 0.
        o_Weight_Stat_Core   = 0.
        o_Weight_Wdg         = 0.       
        # [error case Save]New File Save
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: geometry not valid'.format(OSL_DESIGN_NAME))

# %% [markdown]
# ## [save] New File and  Weight Calc & Export snapshot

# %%
### Export snapshots
    if Pic_Export:
        for screenname in ['Radial', 'Axial', 'StatorWinding']:
            mcApp.SaveScreenToFile(screenname, join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + '_Pic_' + screenname + '.png'))
    mcApp.DisplayScreen('Scripting')   

### Extract active weights
    mcApp.DoWeightCalculation()                                            # Weight calculation
    ex, o_Weight_Mag       = mcApp.GetVariable("Weight_Calc_Magnet")       # Magnet's mass 
    ex, o_Weight_Wdg       = mcApp.GetVariable("Weight_Calc_Copper_Total") # Winding's mass 
    ex, o_Weight_Stat_Core = mcApp.GetVariable("Weight_Calc_Stator_Lam")   # Stator core's mass  
    ex, o_Weight_Rot_Core  = mcApp.GetVariable("Weight_Calc_Rotor_Lam")    # Rotor core's mass
    ex, Weight_Shaft       = mcApp.GetVariable("Weight_Shaft_Active")      # Shaft's mass  
    ex, Weight_Act         = mcApp.GetVariable("Weight_Calc_Total")        # Active mass
    o_Weight_Act           = Weight_Act - Weight_Shaft                     # Shaft's mass retrieved

### Save design
    mcApp.SaveToFile(mot_file_new_path)
    

# %% [markdown]
# ## [Run] (Lab Build -> Lab Analysis(Lab OP study & Duty Cycle) -> Duty Cycle )
#    
# Lab module 
# Shows automatically after assigning options for the saturation & loss models)
# 

# %% [markdown]
# ### Lab: model Build tab### 

# %%
# run_mode='OSL_run'
# run_mode='OSL_setup'

# mcApp = win32com.client.Dispatch("MotorCAD.AppAutomation")  # Launch Motor-CAD application
# mot_file_new_path=r'Z:\Thesis\Optislang_Motorcad\HDEV_Code4\OPD\HDEVCode4Init1.opd\HDEVCODE4\Design0001\HDEV_Model2_Design0001.mot'


# %%
    #mcApp.ClearModelBuild_Lab()  # Clear existing models
    if run_mode in ['IDE_run']:
        mcApp.SetVariable("ModelType_MotorLAB", 2)       # Saturation model type: 1- sigle 2-Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", 0)  # Saturation model: coarse resolution (15 points)    
        mcApp.SetVariable("LossModel_LAB", 0)            # Loss model type: neglect
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", 1)   # Activate saturation model   
    else: 

        mcApp.SetVariable("ModelType_MotorLAB", ModelType_MotorLAB)       # Saturation model type: Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", SatModelPoints_MotorLAB)  # Saturation model: 0 - coarse 1- fine resolution (30 points)  
        mcApp.SetVariable("LossModel_Lab", LossModel_Lab)            # Loss model type: 1-FEA 2 -custom
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", BuildSatModel_MotorLAB)   # Activate saturation model  
        mcApp.SetVariable("BuildLossModel_MotorLAB", BuildLossModel_MotorLAB)  # Activate loss model  
        mcApp.SetVariable("CalcTypeCuLoss_MotorLAB", CalcTypeCuLoss_MotorLAB)  # 0 DC only 1 DC+AC(User) 2 DC+AC (FEA single Point) 3 DC+AC (FEA Map)
        mcApp.SetVariable("IronLossCalc_Lab", IronLossCalc_Lab)          # 0 Neglect 1 OC+SC(User) 2 OC+SC (FEA single Point) 3 (FEA Map)
        mcApp.SetVariable("LabModel_MagnetLoss_Method", LabModel_MagnetLoss_Method) #0 Neglect 1 User Defined 2 OC+SC (FEA single Point) 3 (FEA Map)
        mcApp.setVariable("MagnetLossCalc_Lab",MagnetLossCalc_Lab) # 0 Neglect 1 OC+SC(User) 2 OC+SC (FEA single Point) 3 (FEA Map

    # mcApp.SetVariable("MaxModelCurrent_RMS_MotorLAB", i_BuildLineCurrentRMS)       # Max line current (rms)
    mcApp.SetVariable("MaxModelCurrent_MotorLAB", i_BuildLineCurrentRMS*np.sqrt(2))   # Max line current (peak)
    mcApp.SetVariable('ModelBuildSpeed_MotorLAB', p_Speed_Max)                  # Maximum operating speed
    o_Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3]
    ### Export snapshots
    # if Pic_Export:
    #     for screenname in ['Radial', 'Axial', 'StatorWinding']:
    #         mcApp.SaveScreenToFile(screenname, join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + '_Pic_' + screenname + '.png'))
    # mcApp.DisplayScreen('Scripting')   

### Extract active weights
    mcApp.DoWeightCalculation()                                            # Weight calculation
    ex, o_Weight_Mag       = mcApp.GetVariable("Weight_Calc_Magnet")       # Magnet's mass 
    ex, o_Weight_Wdg       = mcApp.GetVariable("Weight_Calc_Copper_Total") # Winding's mass 
    ex, o_Weight_Stat_Core = mcApp.GetVariable("Weight_Calc_Stator_Lam")   # Stator core's mass  
    ex, o_Weight_Rot_Core  = mcApp.GetVariable("Weight_Calc_Rotor_Lam")    # Rotor core's mass
    ex, Weight_Shaft       = mcApp.GetVariable("Weight_Shaft_Active")      # Shaft's mass  
    ex, Weight_Act         = mcApp.GetVariable("Weight_Calc_Total")        # Active mass
    o_Weight_Act           = Weight_Act - Weight_Shaft                     # Shaft's mass retrieved

### Save design
    mcApp.SaveToFile(mot_file_new_path)
    mcApp.BuildModel_Lab()                                                      # Build activated models
    OP1=['OP1_temp_rise',1700,1200]
    OP2=['OP2_temp_rise',1700,900]
    OP3=['OP3_temp_rise',4000,380]
    M2=['M2','varius']
################################################ Lab: end of Build tab ############################################################### 

# %% [markdown]
# ### Lab :(function Call) Max Torque (Electromagnetic tab)
# [Go to fun_MaxTorque](#fun_maxtorque)
# 

# %%
    o_Maxtorque=fun_MaxTorque(motdata,p_Speed_Peak_Array,i_lineCurrentLabpk,i_TurnLab)   # In [Nm]


# %% [markdown]
# ### Lab :(function Call) BEMF Check (Post Calc)
# [Go to funCalcEMFFundLab](#funcalcemffundlab)
# 

# %%
    o_LineEmfFund=funCalcEMFFundLab(6000,1)

# %% [markdown]
# ### Lab: (function call) peak performance (Duty cycle tab)
# 
# [Go to fun_Find_Ipk_4Trq65C_mk_dat](#fun_Find_Ipk_4Trq65C_mk_dat)\
# [Go to fun_scail_J_by_Turn](#fun_scail_J_by_Turn)
# 
# [Go to fun_OP_temp_contraints](#fun_OP_temp_contraints)
# 
# [Go to fun_TorqueVolumeDensity](#fun_TorqueVolumeDensity)\
# [Go to fun_TorqueWeightDensity](#fun_TorqueWeightDensity)
# 
# [Go to fun_current_density_cal](#fun_current_density_cal)
# 
# 
# 

# %%
   ## Torque check
    o_Op1_ipk,OP1_beta,o_Op1_ShaftToruqe,Op1_after,Op1_current_density,o_Op1_copper_area =fun_Find_Ipk_4Trq65C_mk_dat(OP1,i_TurnLab)
    o_Op1_Jrms=fun_scail_J_by_Turn(Op1_current_density,p_init_Turns_Coil,i_TurnLab)

   ## Temperature rise check
    # o_Op2_ipk,OP2_beta,o_Op2_ShaftToruqe,Op2_after,Op2_current_density,o_Op2_copper_area =fun_Find_Ipk_4Trq65C_mk_dat(OP2,i_TurnLab)
    # o_Op3_ipk,OP3_beta,o_Op3_ShaftToruqe,Op3_after,Op3_current_density,o_Op3_copper_area =fun_Find_Ipk_4Trq65C_mk_dat(OP3,i_TurnLab)

    o_Op2_max_temp,Op2_max_pos,Op2_temp_dic, o_Op2_ipk,Op2_beta,Op2_after,o_Op2_ShaftToruqe, Op2_current_density,o_Op2_copper_area=fun_OP_temp_contraints(OP2,i_TurnLab,OP2thermalCouplingType)
    o_Op3_max_temp,Op3_max_pos,Op3_temp_dic, o_Op3_ipk,Op3_beta,Op3_after,o_Op3_ShaftToruqe, Op3_current_density,o_Op3_copper_area=fun_OP_temp_contraints(OP3,i_TurnLab,OP3thermalCouplingType)

   ## 
    o_TorqueVolumeDensity = fun_TorqueVolumeDensity(o_Maxtorque, o_Active_Volume)   # In [Nm/m3]
    o_TorqueWeightDensity = fun_TorqueWeightDensity(o_Maxtorque,o_Weight_Act)   # In [Nm/kg]
    o_LabCurrentJ,o_Labcopper_area = fun_current_density_cal(i_lineCurrentLabpk,i_TurnLab)

    ## delete for wide design space exploration (the data to be used in surrogate model)
    # if (o_Op1_ShaftToruqe ) < 0:
    #     mcApp.SaveToFile(mot_file_new_path)  # Save design   
    #     mcApp.Quit()                         # Close Motor-CAD
    #     mcApp = 0                            # Reset mcApp variable  
    #     time.sleep(0.5)                      # Frozen for 0.5s
    #     raise Exception('[ERROR] {}: Op1 is fail failed'.format(OSL_DESIGN_NAME))
    # if (o_Op1_ShaftToruqe ) < 0:
    #     mcApp.SaveToFile(mot_file_new_path)  # Save design   
    #     mcApp.Quit()                         # Close Motor-CAD
    #     mcApp = 0                            # Reset mcApp variable  
    #     time.sleep(0.5)                      # Frozen for 0.5s
    #     raise Exception('[ERROR] {}: Op1 is fail failed'.format(OSL_DESIGN_NAME))

#   Raise exception if negative value    
    if (o_Op2_max_temp ) < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}:  Duty Cycle calculation failed'.format(OSL_DESIGN_NAME))
    
    # if (o_Op3_max_temp ) < 0:
    #     mcApp.SaveToFile(mot_file_new_path)  # Save design   
    #     mcApp.Quit()                         # Close Motor-CAD
    #     mcApp = 0                            # Reset mcApp variable  
    #     time.sleep(0.5)                      # Frozen for 0.5s
    #     raise Exception('[ERROR] {}:  Duty Cycle calculation failed'.format(OSL_DESIGN_NAME))


# %% [markdown]
# ### Lab: Driving Duty Cycle (Duty Cycle tab)
# [Go to fun_Driving_Duty](#fun_Driving_Duty)
# 
# [Go to fun_scail_J_by_Turn](#fun_scail_J_by_Turn)
# 

# %%
    M1=['M1_Nosync','M1']
    o_Wh_Loss,o_Wh_Shaft,o_Wh_input=fun_Driving_Duty(M1,i_TurnLab)

        # J(rms) need to be scailed by the scailed number of turns
    o_Op2_Jrms=fun_scail_J_by_Turn(Op2_current_density,p_init_Turns_Coil,i_TurnLab)
    o_Op3_Jrms=fun_scail_J_by_Turn(Op3_current_density,p_init_Turns_Coil,i_TurnLab)

    #   # Raise exception if negative value    
    # if (o_Wh_Loss or o_max_temp  ) < 0:
    #     mcApp.SaveToFile(mot_file_new_path)  # Save design   
    #     mcApp.Quit()                         # Close Motor-CAD
    #     mcApp = 0                            # Reset mcApp variable  
    #     time.sleep(0.5)                      # Frozen for 0.5s
    #     raise Exception('[ERROR] {}: Driving Duty Cycle calculation failed'.format(OSL_DESIGN_NAME))


# %% [markdown]
# ## Close

# %%
## Close Motor-CAD (necessary when running designs in parallel)
    mcApp.SaveToFile(mot_file_new_path)  # Save model
    mcApp.Quit()                         # Close Motor-CAD
    mcApp = 0                            # Reset mcApp variable  
    time.sleep(0.5)                      # Freeze for 0.5s

# %% [markdown]
# ## (else) Responses to be drag and drop during 'OSL_setup' mode 
# 

# %%
else:
  # output initialisation  
    # outputlistOSL=fun_output_define()
    # for i in range(int(len(outputlistOSL)/2)):
    #     exec("%s = %d" % (outputlistOSL[2*i],outputlistOSL[2*i+1]))
    o_LabCurrentJ=0
    # OP point check
    o_Op1_Jrms =0
    o_Op2_Jrms =0
    o_Op3_Jrms =0
    o_Op1_ipk               =0
    o_Op2_ipk               =0
    o_Op3_ipk               =0
    o_Op1_ShaftToruqe       =0
    o_Op2_ShaftToruqe       =0
    o_Op3_ShaftToruqe       =0
    # From Lab data
    o_LineEmfFund   =0
    # Temp rise test (Duty Cycle)
    o_Op2_max_temp          =0
    # o_OP3_max_temp          =0
    # Driving Duty Cycle
    o_Wh_Loss           =0
    o_Wh_Shaft          =0
    o_Wh_input          =0
    # Torque density 
    o_Maxtorque=0
    o_TorqueVolumeDensity =0
    o_TorqueWeightDensity =0
    o_Op2_copper_area   =0
    o_Op1_copper_area   =0
    o_Op3_copper_area   =0
    # Weight
    o_Weight_Act         = 0.
    o_Weight_Mag         = 0.
    o_Weight_Rot_Core    = 0.
    o_Weight_Stat_Core   = 0.
    o_Weight_Wdg         = 0.      
    





