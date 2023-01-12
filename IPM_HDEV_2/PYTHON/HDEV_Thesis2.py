"""
# ----------------------------
# IPM Inverted V_web Templete
# ----------------------------
# MOTOR-CAD : v15.1.7         
# OPTISLANG : v2022 R1 
# PYTHON    : v3.
# ----------------------------
"""
#-----------------------------------------------------      Packages        ----------------------------------------------------

import win32com.client
import os
import numpy as np
import time
from os import getcwd
from os.path import join, dirname, exists
from math import pi, sqrt
from scipy.io import loadmat
import re
# from MotorCAD_Methods import MotorCAD
#------------------------------------------   Running Mode & Working directories     -------------------------------------------

# Modes
#    * OSL-runtime mode : in optiSLang environment, when running a parametric, sensitivity or optimisation study
#    * OSL-setup mode   : in optiSlang environment, when initialising the script and i/o (Python integration node)
#    * IDE-run          : in IDE console, for testing purpose only

# Mode detection
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


# ------------------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------    USER-DEFINED   ---------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------------

#-------------------------------------------------       Functions        ------------------------------------------------------

## Post Work flow
def fun_load_temp_rise_2csvfile(lab_transient_fullpath_w_filename):
    Temp_filename=lab_transient_fullpath_w_filename+'_temp.csv'
    power_filename=lab_transient_fullpath_w_filename+'_power.csv'
    mcApp.ExportResults('Transient', power_filename)
    mcApp.ExportResults('Transient', Temp_filename)
    

def fun_temp_rise_external_dutycyle(path,ext_Duty_Cycle):
    ext_Duty_Cycle_full=join(path,ext_Duty_Cycle,'.dat')
    mcapp.LoadDutyCycle(ext_Duty_Cycle_full)
    mcapp.SetVariable('InitialTransientTemperatureOption',3)
    mcapp.CalculateDutyCycle_Lab()


def fun_rename_matfile_lab_duty(ext_Duty_Cycle):
    ex, motpath=mcApp.GetVariable("CurrentMotFilePath_MotorLAB")
    motpath=re.sub(".mot","",motpath)
    Lab_path=motpath+'/Lab/'
    os.chdir(Lab_path)
    rename_matfile=ext_Duty_Cycle+'_lab_result.mat'
    if os.path.exists('MotorLAB_drivecycledata.mat'):
        os.rename('MotorLAB_drivecycledata.mat',rename_matfile)
        
    
def fun_load_matfile(ext_Duty_Cycle):
    ## This mat file consisted of Ndarray when we are using loadmat function
    ex, motpath=mcApp.GetVariable("CurrentMotFilePath_MotorLAB")
    motpath=re.sub(".mot","",motpath)
    Lab_path=motpath+'/Lab/'
    os.chdir(Lab_path)
    rename_matfile=ext_Duty_Cycle+'_lab_result.mat'
    Mat_File_Data=loadmat(rename_matfile)
    return Mat_File_Data

def fun_check_temp_rise_allcomponent(ext_Duty_Cycle):
    ## init
    init_final_temp=[]
    ## change name of mat file
    fun_rename_matfile_lab_duty(ext_Duty_Cycle)
    ## load mat file
    list_from_mat=fun_load_matfile(ext_Duty_Cycle)
    
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
    
    return dic_init_final_temp, max_temp_key, check_temp





## Simple Calc


def fun_Turn_byAmpT(i_AmpT,i_Line_Current_RMS):
    res = i_AmpT/i_Line_Current_RMS
    return res

def fun_Ipk_beta_by_Trq():
    ex, ipk = mcApp.GetVariable("LabOpPoint_StatorCurrent_Line_Peak")    # Get shaft torque value
    ex, beta = mcApp.GetVariable("LabOpPoint_PhaseAdvance")    # Get shaft torque value
    return ipk, beta

### Geometry


def fun_YtoT(u_YtoT, i_Tooth_Width):
    res = i_Tooth_Width*u_YtoT
    return res

# Machine length
def fun_Machine_Length(p_EndSpace_Height,p_Wdg_Overhang_F, p_Wdg_Overhang_R, i_Active_Length):
    res = i_Active_Length + 2*(p_EndSpace_Height) + p_Wdg_Overhang_F+p_Wdg_Overhang_R
    return res

# Active volume
def fun_Active_Volume(Stator_OD, Active_Length):  
    res = pi*Stator_OD**2/4*Active_Length*1e-9
    return res

# Air pocket
# def fun_Air_Pocket(Mag_Thick, Mag_Clear):
#     res = (Mag_Thick + Mag_Clear)/2 
#     return res

### Performances

# Torque density
def fun_Torque_Density(Torque, Volume):           
    res = Torque/Volume*1e-3
    return res
 
# Stress safety factor
def fun_Stress_Safety(Rotor_Yield, Stress_Max):            
    res = Rotor_Yield/Stress_Max
    return res

#-------------------------------------------------        Inputs        --------------------------------------------------------

### Motor-CAD options
Design_Name     = "HDEV_Thesis2"   # Reference Motor-CAD design
Visible_Opt     = 1.           # Set Motor-CAD visible
Message_Display = 2.           # Display all pop-up messages 
Save_Prompt     = 1.           # Never prompt to save file

### Geometry
p_Pole_Pair            = 6.    # Number of rotor poles
p_Stator_Slots         = 72.   # Number of stator slots

# Predefine Variable

### Stator

#### Absoulute Input (left table)
i_Slot_Corner_Radius    = 1.4     
#i_Tooth_Width  = 6     
p_Tooth_Tip_Depth       = 1    # Tooth tip depth 
p_Tooth_Tip_Angle       = 20   # Tooth tip angle 
#i_Stator_OD             = 400  # Stator outer diameter


#### Ratio(Stator w Hierachy)

p_Slot_Depth_Ratio      =   1   
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

# ### Materials
p_Yield_Rotor  = 460.   # Rotor core yield strength
p_Temp_Wdg_Max = 180.   # Maximum winding temperature
p_Temp_Mag_Max = 140.   # Maximum magnet temperature

### Performance
p_Speed_Max        = 5500.   # Maximum operating speed

### Calculation settings
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
                    
### Input parameters for testing in IDE or initialisation in OSL Python node
if run_mode in ['OSL_setup', 'IDE_run']:
    
    #### Absoulute Input (left table)
    # i_Slot_Corner_Radius    = 1.4     
    # p_Tooth_Tip_Depth       = 1    # Tooth tip depth 
    # p_Tooth_Tip_Angle       = 20   # Tooth tip angle 
    i_Active_Length     = 130.   # Active length
    ### Performance
    i_Line_Current_RMS = 636.3961030678927     # Maximum RMS line current  1000Apk
    i_AmpT_rms=1750                  # Maximum Ampere turn current 2.75T*1000Apk 
    ### Winding
    i_Turns_Coil    = 11  # Number of turns per coil
    i_init_Turns_Coil =11
    i_Tooth_Width           = 6
    i_Stator_OD             = 400  # Stator outer diameter
    #### Ratio(Stator w Hierachy)
    i_Tooth_Width_Ratio                 =0.6
    i_Split_Ratio                       =0.77    
    u_YtoT                              =2.5                                           #ratio user defined YtoT 
                                                                     #get absolute value from ratio
    i_MinBackIronThickness             =fun_YtoT(u_YtoT,i_Tooth_Width)                             
    i_Slot_Op_Ratio        =0.8     

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
    L2_Pole_V_angle          =180 
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


### Real solver run: 'IDE_run' mode or 'OSL_run' mode
if run_mode.endswith('run'):   
    
    
### ----------------------------------------------------------------------------------------------------------------------------
### --------------------------------------------           CALCULATIONS          -----------------------------------------------
### ---------------------------------------------------------------------------------------------------------------------------- 
    
### --------------------------------------------    Pre-calculations in PYTHON    ----------------------------------------------
       
### Geometry parameters
    Machine_Length = fun_Machine_Length(p_EndSpace_Height, p_Wdg_Overhang, i_Active_Length)
    # Air_Pocket     = fun_Air_Pocket(i_Mag_Thick, p_Mag_Clear)
    
### ---------------------------------------------------      MOTOR-CAD     -----------------------------------------------------

### Load reference Motor-CAD file
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
    
### EMag context
    mcApp.ShowMagneticContext()          
    mcApp.DisplayScreen('Scripting')     # Switch to a tab where no parameter is adjusted 
    
### Change to ratio mode to edit the geometry
    mcApp.SetVariable('GeometryParameterisation', 1)   # Ratio mode in Motor-CAD 

### Assign geometry parameters
    mcApp.SetVariable('Slot_Number', p_Stator_Slots)                           # Stator slots

    ### Stator
    #### Absoulute Input (left table)
    mcApp.SetVariable('Stator_Lam_Dia'      , i_Stator_OD)                        # Stator OD 
    mcApp.SetVariable('Slot_Corner_Radius'  , i_Slot_Corner_Radius)               # Slot_Corner_Radius      
    mcApp.SetVariable('Tooth_Tip_Depth'     , p_Tooth_Tip_Depth)                  # Tooth tip depth
    mcApp.SetVariable('Tooth_Tip_Angle'     , p_Tooth_Tip_Angle)                  # Tooth tip angle
    Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3]


    #### Ratio
    #u_YtoT                              =2.5                                         #ratio user defined YtoT 
    ex, i_Tooth_Width                   = mcApp.GetVariable("Tooth_Width")           # Absolute Tooth_Width
    # i_Tooth_Width  = 6                                                                     #get absolute value from ratio
    i_MinBackIronThickness              = fun_YtoT(u_YtoT,i_Tooth_Width) 

    mcApp.SetVariable("Ratio_SlotDepth_ParallelTooth"           , p_Slot_Depth_Ratio     )        #Ratio_SlotDepth_ParallelTooth" )   %% Fixed     
    mcApp.SetVariable('MinBackIronThickness'                    , i_MinBackIronThickness )        #Abosolute be user-defined with Y to T ratio     
    mcApp.SetVariable("Ratio_ToothWidth"                        , i_Tooth_Width_Ratio    )        #Ratio_ToothWidth" )      
    mcApp.SetVariable("Ratio_SlotOpening_ParallelTooth"         , i_Slot_Op_Ratio        )        #Ratio_SlotOpening_ParallelTooth" )      
    
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

### Check the geometry is valid
    success = mcApp.CheckIfGeometryIsValid(0)
    if success == 0:
        # If not valid, generate zero outputs instead of getting an error message in optiSLang
        o_Cont_Torque_1700rpm  = 0.  
        o_Cont_Torque_4krpm  = 0.

        o_Peak_Torque_1700rpm = 0.
        o_Peak_Power_1700rpm   = 0
        o_Wh_Loss           =0
        o_Wh_Shaft          =0
        o_Wh_input          =0
        o_Current_Density    = 0.
        o_Torque_Density     = 0.
        o_LabPeak_IPeak     =0.
        #o_Torque_Ripples     = 0.
        o_Weight_Act         = 0.
        o_Weight_Mag         = 0.
        o_Weight_Rot_Core    = 0.
        o_Weight_Stat_Core   = 0.
        o_Weight_Wdg         = 0.
        #o_Stress_Safety      = 0.
        if run_mode in ['OSL_run']:
            o_Sig_Peak_Torque    = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
            o_Sig_Peak_Power     = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
            # o_Sig_Torque_Ripples = list_list_2_variant_signal([[0]*(p_Torque_Pts+1)], [0]*(p_Torque_Pts+1)) 
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: geometry not valid'.format(OSL_DESIGN_NAME))
    
### Assign winding parameters
    mcApp.SetVariable('WindingLayers', p_Coils_Slot)                # Coils passing through a slot
    mcApp.SetVariable('MagTurnsConductor', i_init_Turns_Coil)            # Turns per coil
    mcApp.SetVariable('ParallelPaths', p_Parallel_Path)             # Parallel paths per phase
    mcApp.SetVariable('RequestedGrossSlotFillFactor', p_Slot_Fill)  # Slot fill factor
        
### Assign initial calculation settings
    mcApp.SetVariable("BackEMFCalculation", False)               # OC calculations deactivated
    mcApp.SetVariable("CoggingTorqueCalculation", False)         # Cogging torque calculation deactivated
    mcApp.SetVariable("TorqueCalculation", False)                # Torque calculations deactivated
    mcApp.SetVariable("TorqueSpeedCalculation", False)           # Torque speed curve calculation deactivated
    mcApp.SetVariable("DemagnetizationCalc", False)              # Demagnetisation test deactivated
    mcApp.SetVariable("InductanceCalc", False)                   # Inductance calculation deactivated
    mcApp.SetVariable("BPMShortCircuitCalc", False)              # Transient short circuit calculation deactivated
    mcApp.SetVariable("ElectromagneticForcesCalc_OC", False)     # Maxwell forces calculation deactivated (OC)
    mcApp.SetVariable("ElectromagneticForcesCalc_Load", False)   # Maxwell forces calculation deactivated (OL)
    mcApp.SetVariable("MagneticSolver", 1)                       # Transient calculation enabled
    mcApp.SetVariable("Lab_Threads_Enabled", True)               # Threading option for lab models enabled
    
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
    
### Lab module
### Shows automatically after assigning options for the saturation & loss models)

### Lab: model Build tab 
    #mcApp.ClearModelBuild_Lab()  # Clear existing models
    if run_mode in ['IDE_run']:
        mcApp.SetVariable("ModelType_MotorLAB", 2)       # Saturation model type: 1- sigle 2-Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", 0)  # Saturation model: coarse resolution (15 points)    
        mcApp.SetVariable("LossModel_LAB", 0)            # Loss model type: neglect
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", 1)   # Activate saturation model   
    else: 
        mcApp.SetVariable("ModelType_MotorLAB", 1)       # Saturation model type: Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", 0)  # Saturation model: 0 - coarse 1- fine resolution (30 points)  
        mcApp.SetVariable("LossModel_Lab", 2)            # Loss model type: 1-FEA 2 -custom
        
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", 1)   # Activate saturation model  
        mcApp.SetVariable("BuildLossModel_MotorLAB", 1)  # Activate loss model  
        mcApp.SetVariable("CalcTypeCuLoss_MotorLAB", 2)  # 0 DC only 1 DC+AC(User) 2 DC+AC (FEA single Point) 3 DC+AC (FEA Map)
        mcApp.SetVariable("IronLossCalc_Lab", 2)          # 0 Neglect 1 OC+SC(User) 2 OC+SC (FEA single Point) 3 (FEA Map)
        mcApp.SetVariable("LabModel_MagnetLoss_Method", 2) #0 Neglect 1 User Defined 2 OC+SC (FEA single Point) 3 (FEA Map)

    # mcApp.SetVariable("MaxModelCurrent_RMS_MotorLAB", i_Line_Current_RMS)       # Max line current (rms)
    mcApp.SetVariable("MaxModelCurrent_MotorLAB", i_Line_Current_RMS*np.sqrt(2))   # Max line current (peak)
    mcApp.SetVariable('ModelBuildSpeed_MotorLAB', p_Speed_Max)                  # Maximum operating speed
    Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3]

    mcApp.BuildModel_Lab()                                                      # Build activated models
              
### Lab: peak performance
    i_Turns_Coil    = fun_Turn_byAmpT(i_AmpT_rms,i_Line_Current_RMS)*p_Parallel_Path   # Number of turns per coil
    mcApp.SetVariable('TurnsCalc_MotorLAB', i_Turns_Coil)            # Turns per coil
    mcApp.SetVariable("LabThermalCoupling", 0)                         # Coupling with Thermal
    mcApp.SetVariable("OpPointSpec_MotorLAB", 0)                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("SpeedDemand_MotorLAB", p_Speed_Peak_Array[0])                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("TorqueDemand_MotorLAB", 1300)                       # 0- Torque 4-Max temperature definition  
    mcApp.CalculateOperatingPoint_Lab()                             # Operating point calculation
    
    # ex, LabOpPoint_ShaftTorque = mcApp.GetVariable("LabOpPoint_ShaftTorque")    # Get shaft torque value
    o_LabPeak_IPeak, OLabPeak_beta=fun_Ipk_beta_by_Trq()
    
  # Settings
    mcApp.SetVariable("OperatingMode_Lab", 0)                      # Motoring mode
    mcApp.SetVariable("EmagneticCalcType_Lab", 0)                  # 0- Peak performance 1- Effi_map
    mcApp.SetVariable('SpeedMax_MotorLAB', p_Speed_Max)            # Maximum speed
    mcApp.SetVariable('Speedinc_MotorLAB', p_Speed_Lab_Step)       # Speed step
    mcApp.SetVariable('Imax_RMS_MotorLAB', o_LabPeak_IPeak)     # Max line current (rms)
    mcApp.SetVariable('Imax_MotorLAB', o_LabPeak_IPeak*sqrt(2)) # Max line current (peak) 
    
  # Calculation & Data management
    mcApp.CalculateMagnetic_Lab()                                 # Run calculation
    Mat_File_Name     = 'MotorLAB_elecdata.mat'                   # *.mat file automatically generated by Motor-CAD
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
    Peak_Torque_Array = np.zeros(len(p_Speed_Peak_Array))  # Peak Torque array initialisation
    for i in range(len(p_Speed_Peak_Array)):
        Ind_Speed = (np.abs(Mat_File_Speed - p_Speed_Peak_Array[i])).argmin()   # Index corresponding to the required speed
        Peak_Power_Array[i]  = Mat_File_Power[Ind_Speed]                        # Peak power at the required speed 
        Peak_Torque_Array[i] = Mat_File_Torque[Ind_Speed]                       # Peak torque at the required speed 
    o_Peak_Torque_1700rpm    = Peak_Torque_Array[0]      # In [Nm] 
    o_Peak_Power_1700krpm      = Peak_Power_Array[0]       # In [kW]    
  
  # Raise exception if wrong performance data
    if (o_Peak_Torque_1700rpm or o_Peak_Power_1700krpm) < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: peak performance calculation failed'.format(OSL_DESIGN_NAME))

  # Key performance indicators
    o_Torque_Density = fun_Torque_Density(o_Peak_Torque_1700rpm, Active_Volume)  # Torque density [Nm/L]

  # Signals to be read in OSL 
    if run_mode in ['OSL_run']:
        Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3]
        List_Speed        = Mat_File_Speed.tolist()     # Convert to list 
        List_Peak_Torque  = Mat_File_Torque.tolist()
        List_Peak_Power   = Mat_File_Power.tolist()
        o_Sig_Peak_Torque = list_list_2_variant_signal([List_Peak_Torque], List_Speed)
        o_Sig_Peak_Power  = list_list_2_variant_signal([List_Peak_Power], List_Speed)
    
    
    mcApp.SetVariable("LabThermalCoupling", 0)                         # Coupling with Thermal
    mcApp.SetVariable("OpPointSpec_MotorLAB", 0)                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("SpeedDemand_MotorLAB", p_Speed_Peak_Array[0])                       # 0- Torque 4-Max temperature definition
    mcApp.SetVariable("TorqueDemand_MotorLAB", 900)                       # 0- Torque 4-Max temperature definition  
    mcApp.CalculateOperatingPoint_Lab()                             # Operating point calculation
    
    # ex, LabOpPoint_ShaftTorque = mcApp.GetVariable("LabOpPoint_ShaftTorque")    # Get shaft torque value
    o_900NmLabPeak_IPeak=fun_Ipk_by_Trq()
    o_900NmLab_Beta=fun_beta_by_Trq()


### Lab: Energy Consumption over drive cycle

  # Settings
    # mcApp.SetVariable('N_d_MotorLAB', i_Gear_Ratio)                     # Gear Ratio

    mcApp.SetVariable("DutyCycleType_Lab", 0)                 # Automotive drive cycle
    # mcApp.SetVariable("DrivCycle_MotorLAB", "WLTP Class 3")   # WLTP3 drive cycle
    mcApp.SetVariable("LabThermalCoupling_DutyCycle", 0)      # No coupling with Thermal
    
  # Calculation & Post processing
    mcApp.CalculateDutyCycle_Lab()                                                # Run calculation
    #ex, o_WLTP3_Eff = mcApp.GetVariable("DutyCycleAverageEfficiency_EnergyUse")   # Get efficiency value 
    ex, o_Wh_Loss = mcApp.GetVariable("DutyCycleTotalLoss")   # Get efficiency value 
    ex, o_Wh_Shaft = mcApp.GetVariable("DutyCycleTotalEnergy_Shaft_Output")   # Get efficiency value 
    ex, o_Wh_input = mcApp.GetVariable("DutyCycleTotalEnergy_Electrical_Input")   # Get efficiency value 
      
  # Raise exception if wrong performance data
    if o_Wh_Loss < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: drive cycle performance calculation failed'.format(OSL_DESIGN_NAME))

### Lab: continuous performance Thermal

  # Settings Thermal
    # mcApp.SetVariable("LabThermalCoupling", 1)                         # Coupling with Thermal
    # mcApp.SetVariable("OpPointSpec_MotorLAB", 2)                       # Max temperature definition
    # mcApp.SetVariable("StatorTempDemand_Lab", p_Temp_Wdg_Max)          # Max winding temperature
    # mcApp.SetVariable("RotorTempDemand_Lab", p_Temp_Mag_Max)           # Max magnet temperature
    # mcApp.SetVariable("ThermMaxCurrentLim_MotorLAB", 0)                # Useful?
    # mcApp.SetVariable("Iest_MotorLAB", i_Line_Current_RMS*sqrt(2)*0.3) # Initial line current (rms)
    # mcApp.SetVariable("Iest_RMS_MotorLAB", i_Line_Current_RMS*0.3)     # Initial line current (peak)        
    
  # Calculation & Post processing
    # Cont_Torque_Array  = np.zeros(len(p_Speed_Cont_Array))              # Contiuous torque array initialisation
    # for i in range(len(p_Speed_Cont_Array)):                            # Calculation between Lab and Thermal
    #     mcApp.SetVariable("SpeedDemand_MotorLAB", p_Speed_Cont_Array[i])# Set operating speed
    #     mcApp.CalculateOperatingPoint_Lab()                             # Operating point calculation          
    #     ex, ContTorque = mcApp.GetVariable("LabOpPoint_ShaftTorque")    # Get shaft torque value
    #     Cont_Torque_Array[i] = ContTorque
    # o_Cont_Torque_1700rpm = Cont_Torque_Array[0]                          # Assigning array variables
    # o_Cont_Torque_4krpm = Cont_Torque_Array[1]

#   # Raise exception if negative value    
#     if (o_Cont_Torque_1700rpm or o_Cont_Torque_4krpm) < 0:
#         mcApp.SaveToFile(mot_file_new_path)  # Save design   
#         mcApp.Quit()                         # Close Motor-CAD
#         mcApp = 0                            # Reset mcApp variable  
#         time.sleep(0.5)                      # Frozen for 0.5s
#         raise Exception('[ERROR] {}: continuous performance calculation failed'.format(OSL_DESIGN_NAME))
### ---------------------------------------------------      SCREENSHOTS   -----------------------------------------------------

### Extract operating point for torque ripple calculation in EMag
    #mcApp.SetVariable("OpPointSpec_MotorLAB", 1)                             # Max Current definition    
    #mcApp.SetVariable("StatorCurrentDemand_Lab", i_Line_Current_RMS*sqrt(2)) # Line current (rms)
    #mcApp.SetVariable("StatorCurrentDemand_RMS_Lab", i_Line_Current_RMS)     # Line current (peak)
    #mcApp.SetVariable("SpeedDemand_MotorLAB", 500)                           # Operating speed        
    #mcApp.SetVariable("LabThermalCoupling", 0)                               # No coupling with Thermal
    #mcApp.SetVariable("LabMagneticCoupling", 1)                              # Coupling with EMag
    #mcApp.CalculateOperatingPoint_Lab()                                      # Export to EMag
  
### Current density
    o_Current_Density = mcApp.GetVariable('ArmatureConductorCurrentDensity')
    o_Current_Density = o_Current_Density[1]
    ex, i_Tooth_Width                   = mcApp.GetVariable("Tooth_Width")        # Absolute Tooth_Width

## Close Motor-CAD (necessary when running designs in parallel)
    mcApp.SaveToFile(mot_file_new_path)  # Save model
    mcApp.Quit()                         # Close Motor-CAD
    mcApp = 0                            # Reset mcApp variable  
    time.sleep(0.5)                      # Freeze for 0.5s
    
### ----------------------------------------------      INITIALISATION (END)     ------------------------------------------------

### Responses to be drag and drop during 'OSL_setup' mode 
else:
    
  # Scalars
    o_Cont_Torque_1700rpm  = 0.
    o_Cont_Torque_4krpm  = 0.
    o_Peak_Power_1700rpm   = 0.
    o_Peak_Torque_1700rpm = 0.
    #o_WLTP3_Eff          = 0.
    o_Current_Density    = 0.
    o_Torque_Density     = 0.
    # o_Torque_Ripples     = 0.
    o_Weight_Act         = 0.
    o_Weight_Mag         = 0.
    o_Weight_Rot_Core    = 0.
    o_Weight_Stat_Core   = 0.
    o_Weight_Wdg         = 0.
    # o_Stress_Safety      = 0.
    o_Wh_Loss           =0
    o_Wh_Shaft          =0
    o_Wh_input          =0
    o_LabPeak_IPeak = 0
  # Signals
    o_Sig_Peak_Torque    = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
    o_Sig_Peak_Power     = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
    # o_Sig_Torque_Ripples = list_list_2_variant_signal([[0]*(p_Torque_Pts+1)], [0]*(p_Torque_Pts+1))
