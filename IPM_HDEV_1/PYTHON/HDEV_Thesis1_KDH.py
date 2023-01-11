"""
# ----------------------------
# IPM Inverted V_web Templete
# ----------------------------
# MOTOR-CAD : v15.1.7         
# OPTISLANG : v2022 R1 
# PYTHON    : v3.
# ----------------------------
"""

### ----------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------             INITALIZATION          -----------------------------------------------
### ----------------------------------------------------------------------------------------------------------------------------

# The script works using the following folder structure:
#     * OPD folder: contains the optiSLang files            e.g. path: "MyProjectPath\OPD..."
#     * MCAD folder: contains the reference Motor-CAD file  e.g. path: "MyProjectPath\MCAD..."
#     * PYTHON folder: contains the scripts                 e.g. path: "MyProjectPath\PYTHON..."

#-----------------------------------------------------      Packages        ----------------------------------------------------

import win32com.client
import os
import numpy as np
import time
from os import getcwd
from os.path import join, dirname, exists
from math import pi, sqrt
from scipy.io import loadmat
import matplotlib
import matplotlib.pyplot as plt
from scipy import io
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

### Geometry

# Machine length
def fun_Machine_Length(p_EndSpace_Height, p_Wdg_Overhang, i_Active_Length):
    res = i_Active_Length + 2*(p_EndSpace_Height + p_Wdg_Overhang)
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
Design_Name     = "HDEV_Thesis1"   # Reference Motor-CAD design
Visible_Opt     = 1.           # Set Motor-CAD visible
Message_Display = 2.           # Display all pop-up messages 
Save_Prompt     = 1.           # Never prompt to save file

### Geometry
p_Pole_Pair            = 6    # Number of rotor poles
p_Stator_Slots         = 72   # Number of stator slots

# Predefine Variable

### Stator

#### Absoulute Input (left table)
i_Slot_Corner_Radius    = 1.4     
p_Tooth_Tip_Depth       = 1    # Tooth tip depth 
p_Tooth_Tip_Angle       = 20   # Tooth tip angle 
# i_Stator_OD             = 400  # Stator outer diameter


#### Ratio(Stator w Hierachy)

# ### Rotor 
# #### Absoulute Input (left table)
#     # Notch_Depth=0 
# Magnet_Layers=2 

# #### Ratio  (Rotor w Hierachy)

p_Airgap_Mecha         = 1.    # Mechanical airgap
# p_Mag_Clear            = 0   # Magnet clearance

### other

p_EndSpace_Height      = 15.   # Space between winding ends and caps
p_Wdg_Overhang         = 20.   # Winding overhang height

### Winding
p_Coils_Slot    = 1.    # Number of coils going through each slot
p_Parallel_Path = 4.    # Number of parallel paths per phase
p_Slot_Fill     = 0.5325   # Copper slot fill factor

# ### Materials
# p_Yield_Rotor  = 460.   # Rotor core yield strength
# p_Temp_Wdg_Max = 180.   # Maximum winding temperature
# p_Temp_Mag_Max = 140.   # Maximum magnet temperature

### Performance
p_Speed_Max        = 5000.   # Maximum operating speed

### Calculation settings
p_Speed_Lab_Step    = 100.                              # Speed step used in Lab
p_Speed_Peak_Array  = np.array([500., 3000., 6000.])    # Speeds for peak performance calculation 
p_Speed_Cont_Array  = np.array([1000., 5000.])          # Speeds for continuous performance calculation
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
    i_Stator_OD             = 400  # Stator outer diameter

    #### Ratio(Stator w Hierachy)
    i_Slot_Depth_Ratio      =   1   
    i_Tooth_Width_Ratio                 =0.6
    i_Split_Ratio                       =0.77    
    #u_YtoT                              =2.5                                           #ratio user defined YtoT 
    # ex, i_Tooth_Width                   = mcApp.GetVariable("Tooth_Width")        # Absolute Tooth_Width
    # i_Tooth_Width  = 6                                                                     #get absolute value from ratio
    # i_MinBackIronThickness             =u_YtoT*i_Tooth_Width                           
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

    i_Active_Length     = 130.   # Active length
    ### Winding
    i_Turns_Coil    = 11.   # Number of turns per coil

    ### Performance
    i_Line_Current_RMS = 560.    # Maximum RMS line current
    ### Duty Cycle Study
    i_Gear_Ratio = 7.
### Real solver run: 'IDE_run' mode or 'OSL_run' mode
if run_mode.endswith('run'):   
    
    
### ----------------------------------------------------------------------------------------------------------------------------
### --------------------------------------------           CALCULATIONS          -----------------------------------------------
### ---------------------------------------------------------------------------------------------------------------------------- 
    
### --------------------------------------------    Pre-calculations in PYTHON    ----------------------------------------------
       
### Geometry parameters
    Machine_Length = fun_Machine_Length(p_EndSpace_Height, p_Wdg_Overhang, i_Active_Length)
    # Air_Pocket     = fun_Air_Pocket(i_Mag_Thick, p_Mag_Clear)
    Active_Volume  = fun_Active_Volume(i_Stator_OD, i_Active_Length)     # In [m3]
    
### ---------------------------------------------------      MOTOR-CAD     -----------------------------------------------------

### Load reference Motor-CAD file
    mcApp = win32com.client.Dispatch("MotorCAD.AppAutomation")  # Launch Motor-CAD application
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
 

    #### Ratio
    # u_YtoT                              =2.5                                          #ratio user defined YtoT 
    # ex, i_Tooth_Width                   = mcApp.GetVariable("Tooth_Width")           # Absolute Tooth_Width
    # i_Tooth_Width  = 6                                                                     #get absolute value from ratio
    # i_MinBackIronThickness              =u_YtoT*i_Tooth_Width 
    mcApp.SetVariable("Ratio_SlotDepth_ParallelTooth"           , i_Slot_Depth_Ratio     )        #Ratio_SlotDepth_ParallelTooth" )   %% Fixed  
   
   # mcApp.SetVariable('MinBackIronThickness'                    , i_MinBackIronThickness )        #Abosolute be user-defined with Y to T ratio     
    mcApp.SetVariable("Ratio_ToothWidth"                        , i_Tooth_Width_Ratio    )        #Ratio_ToothWidth" )      
    mcApp.SetVariable("Ratio_SlotOpening_ParallelTooth"         , i_Slot_Op_Ratio        )        #Ratio_SlotOpening_ParallelTooth" )      
    
    ### Rotor 
    #### Absoulute Input (left table)
    mcApp.SetArrayVariable("MagnetThickness_Array"                , 0, L1_Magnet_Thickness)            # Layer 1 Magnet thickness
    mcApp.SetArrayVariable("MagnetThickness_Array"                , 1, L2_Magnet_Thickness)            # Layer 2 Magnet thickness
                    
    mcApp.SetArrayVariable('BridgeThickness_Array'                , 0, L1_Bridge_Thickness)            # Layer 1 Bridge thickness 
    mcApp.SetArrayVariable('BridgeThickness_Array'                , 1, L2_Bridge_Thickness)            # Layer 2 Bridge thickness 
        
    mcApp.SetArrayVariable("PoleVAngle_Array"                     , 0, L1_Pole_V_angle)                # Layer 1 V-shape layer angle
    mcApp.SetArrayVariable("PoleVAngle_Array"                     , 1, L2_Pole_V_angle)                # Layer 2 V-shape layer angle

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
    mcApp.SetVariable('EWdg_Overhang_[R]', p_Wdg_Overhang)                     # End winding overhang (rear)
    mcApp.SetVariable('EWdg_Overhang_[F]', p_Wdg_Overhang)                     #                      (front)

### Check the geometry is valid
    success = mcApp.CheckIfGeometryIsValid(0)
    if success == 0:
        # If not valid, generate zero outputs instead of getting an error message in optiSLang
        # o_Cont_Torque_1krpm  = 0.  
        # o_Cont_Torque_4krpm  = 0.
        o_Peak_Power_1krpm   = 0.
        o_Peak_Power_4krpm    = 0.
        o_Peak_Torque_1700rpm = 0.
        o_WLTP3_Eff          = 0.
        # o_Current_Density    = 0.
        o_Torque_Density     = 0.
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
            o_Sig_Torque_Ripples = list_list_2_variant_signal([[0]*(p_Torque_Pts+1)], [0]*(p_Torque_Pts+1)) 
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: geometry not valid'.format(OSL_DESIGN_NAME))
    
### Assign winding parameters
    mcApp.SetVariable('WindingLayers', p_Coils_Slot)                # Coils passing through a slot
    mcApp.SetVariable('MagTurnsConductor', i_Turns_Coil)            # Turns per coil
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
    mcApp.SetVariable("MagneticSolver", 0)                       # Transient calculation enabled
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
    # mcApp.ClearModelBuild_Lab()  # Clear existing models
    if run_mode in ['IDE_run']:
        mcApp.SetVariable("ModelType_MotorLAB", 2)       # Saturation model type: Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", 0)  # Saturation model: coarse resolution (15 points)    
        mcApp.SetVariable("LossModel_LAB", 0)            # Loss model type: neglect
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", 1)   # Activate saturation model   
    else: 
        mcApp.SetVariable("ModelType_MotorLAB", 2)       # Saturation model type: Full Cycle
        mcApp.SetVariable("SatModelPoints_MotorLAB", 1)  # Saturation model: fine resolution (30 points)  
        mcApp.SetVariable("LossModel_Lab", 2)            # Loss model type: FEA
        mcApp.SetMotorLABContext()                       # Lab context
        mcApp.SetVariable("BuildSatModel_MotorLAB", 1)   # Activate saturation model  
        mcApp.SetVariable("BuildLossModel_MotorLAB", 1)  # Activate loss model  

    mcApp.SetVariable("MaxModelCurrent_RMS_MotorLAB", i_Line_Current_RMS)       # Max line current (rms)
    mcApp.SetVariable("MaxModelCurrent_MotorLAB", i_Line_Current_RMS*sqrt(2))   # Max line current (peak)
    mcApp.SetVariable('ModelBuildSpeed_MotorLAB', p_Speed_Max)                  # Maximum operating speed
    mcApp.BuildModel_Lab()                                                      # Build activated models
              
### Lab: peak performance

  # Settings
    mcApp.SetVariable("OperatingMode_Lab", 0)                      # Motoring mode
    mcApp.SetVariable("EmagneticCalcType_Lab", 0)                  # Peak performance
    mcApp.SetVariable('SpeedMax_MotorLAB', p_Speed_Max)            # Maximum speed
    mcApp.SetVariable('Speedinc_MotorLAB', p_Speed_Lab_Step)       # Speed step
    mcApp.SetVariable('Imax_RMS_MotorLAB', i_Line_Current_RMS)     # Max line current (rms)
    mcApp.SetVariable('Imax_MotorLAB', i_Line_Current_RMS*sqrt(2)) # Max line current (peak) 
    
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
    o_Peak_Power_1krpm      = Peak_Power_Array[1]       # In [kW]    
    o_Peak_Power_4krpm       = Peak_Power_Array[2]       # In [kW]
  
  # Raise exception if wrong performance data
    if (o_Peak_Torque_1700rpm or o_Peak_Power_1krpm or o_Peak_Power_4krpm) < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: peak performance calculation failed'.format(OSL_DESIGN_NAME))

  # Key performance indicators
    o_Torque_Density = fun_Torque_Density(o_Peak_Torque_1700rpm, Active_Volume)  # Torque density [Nm/L]

  # Signals to be read in OSL 
    if run_mode in ['OSL_run']:
        List_Speed        = Mat_File_Speed.tolist()     # Convert to list 
        List_Peak_Torque  = Mat_File_Torque.tolist()
        List_Peak_Power   = Mat_File_Power.tolist()
        o_Sig_Peak_Torque = list_list_2_variant_signal([List_Peak_Torque], List_Speed)
        o_Sig_Peak_Power  = list_list_2_variant_signal([List_Peak_Power], List_Speed)
  
### Lab: efficiency over WLTP-3 drive cycle

  # Settings
    mcApp.SetVariable('N_d_MotorLAB', i_Gear_Ratio)                     # Gear Ratio

    mcApp.SetVariable("DutyCycleType_Lab", 1)                 # Automotive drive cycle
    mcApp.SetVariable("DrivCycle_MotorLAB", "WLTP Class 3")   # WLTP3 drive cycle
    mcApp.SetVariable("LabThermalCoupling_DutyCycle", 1)      # No coupling with Thermal
    
  # Calculation & Post processing
    mcApp.CalculateDutyCycle_Lab()                                                # Run calculation
    ex, o_WLTP3_Eff = mcApp.GetVariable("DutyCycleAverageEfficiency_EnergyUse")   # Get efficiency value 
    ex,    o_DutyCycleTotalEnergy_Electrical_Input   =        mcApp.GetVariable("DutyCycleTotalEnergy_Electrical_Input")           
    ex,    o_DutyCycleTotalEnergy_Shaft_Output       =        mcApp.GetVariable("DutyCycleTotalEnergy_Shaft_Output")
    ex,    o_DutyCycleTotalEnergy_Electrical_Output  =        mcApp.GetVariable("DutyCycleTotalEnergy_Electrical_Output")
    ex,    o_DutyCycleTotalEnergy_Shaft_Input        =        mcApp.GetVariable("DutyCycleTotalEnergy_Shaft_Input")



  # Raise exception if wrong performance data
    if o_WLTP3_Eff < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: drive cycle performance calculation failed'.format(OSL_DESIGN_NAME))


# ### ---------------------------------------------------      SCREENSHOTS   -----------------------------------------------------

### Extract operating point for torque ripple calculation in EMag
    #mcApp.SetVariable("OpPointSpec_MotorLAB", 1)                             # Max Current definition    
    #mcApp.SetVariable("StatorCurrentDemand_Lab", i_Line_Current_RMS*sqrt(2)) # Line current (rms)
    #mcApp.SetVariable("StatorCurrentDemand_RMS_Lab", i_Line_Current_RMS)     # Line current (peak)
    #mcApp.SetVariable("SpeedDemand_MotorLAB", 500)                           # Operating speed        
    #mcApp.SetVariable("LabThermalCoupling", 0)                               # No coupling with Thermal
    #mcApp.SetVariable("LabMagneticCoupling", 1)                              # Coupling with EMag
    #mcApp.CalculateOperatingPoint_Lab()                                      # Export to EMag
  
### EMag context
    mcApp.ShowMagneticContext()       # Back to EMag for torque ripples calculation
    mcApp.DisplayScreen('Scripting') 
    
### Cogging Calculation
    p_CoggingPointsPerCycle  =30    
    gcd_ps=np.gcd(2*p_Pole_Pair,p_Stator_Slots)
    cogging_period=360*gcd_ps/(2*p_Pole_Pair*p_Stator_Slots)
    cogging_period # in mech [deg]
    Rotation_angle_perpoint=cogging_period/p_CoggingPointsPerCycle
    Rotation_angle_perpoint # in [deg]
    N = 1
    Airgap_points1= N*360/Rotation_angle_perpoint 
    Airgap_points2= N*2*p_Pole_Pair*p_Stator_Slots*p_CoggingPointsPerCycle/gcd_ps   
    mcApp.SetVariable("CoggingTorqueCalculation", True)            # Torque calculation
    mcApp.SetVariable("CoggingPointsPerCycle", p_CoggingPointsPerCycle)                  # Reduced multi-static solver

### BEMF calculation

  # Settings
    simu_Period=360/p_Pole_Pair
    simu_Period                 # in mech [deg] for one Elec period
    p_BackEMFPointsPerCycle = 120
    Rotation_angle_BEMF_perpoint=simu_Period/p_BackEMFPointsPerCycle
    Rotation_angle_BEMF_perpoint # in mech [deg]
    Airgap_points_EMF1= N*360/Rotation_angle_BEMF_perpoint 
    Airgap_points_EMF2= N*p_Pole_Pair*p_BackEMFPointsPerCycle

    mcApp.SetVariable("BackEMFCalculation", True)            # Torque calculation
    mcApp.SetVariable("AirgapMeshPoints_layers", Airgap_points1)      # Number of mesh points in the airgap
    mcApp.SetVariable("AirgapMeshPoints_mesh", Airgap_points1)        # Number of mesh points at the airgap surface
    mcApp.SetVariable("MagneticSolver", 0)                  # Reduced multi-static solver
    mcApp.SetVariable("ShaftSpeed", 5000)                  # Reduced multi-static solver
    mcApp.SetVariable("BackEMFPointsPerCycle", p_CoggingPointsPerCycle)                  # Reduced multi-static solver

  # Calculation & Post processing
    mcApp.ShowMagneticContext()          
    mcApp.DisplayScreen('Scripting')     # Switch to a tab where no parameter is adjusted 

  # 
    ex, CoggingPointsPerCycle       = mcApp.GetVariable("CoggingPointsPerCycle")       # Magnet's mass 
    ex, CoggingNumberCycles         = mcApp.GetVariable("CoggingNumberCycles")       # Magnet's mass 
    ex, BackEMFPointsPerCycle       = mcApp.GetVariable("BackEMFPointsPerCycle")       # Magnet's mass 
    ex, BackEMFNumberCycles         = mcApp.GetVariable("BackEMFNumberCycles")       # Magnet's mass      
    NumBackEMFPoints = (BackEMFPointsPerCycle * BackEMFNumberCycles) + 1;
    
    RotorPosition = np.zeros(NumBackEMFPoints);
    
    # CoggingTorqueVW = np.zeros(NumTorquePoints);
    
    BackEMFPh1      = np.zeros(NumBackEMFPoints);
    BackEMFLineToLine12 =  np.zeros(NumBackEMFPoints);
  # Cogging Export 
    for loop in range(NumBackEMFPoints-1):
        # ex, x, cogg =mcApp.GetMagneticGraphPoint('CoggingTorqueVW', loop)
        ex, x, BEMFPh =mcApp.GetMagneticGraphPoint('BackEMFPh1', loop)
        ex, x, BEMFline =mcApp.GetMagneticGraphPoint('BackEMFLineToLine12', loop)
        RotorPosition[loop+1]=x
        # CoggingTorqueVW[loop+1]=cogg
        BackEMFPh1[loop+1]=BEMFPh
        BackEMFLineToLine12[loop+1]=BEMFline

    fig, ax = plt.subplots()
    ax.plot(RotorPosition, BackEMFPh1)
    ax.set(xlabel='Rotor Position (Mdeg)', ylabel='BackEMFPh1 [V]',
    title='BackEMFPh1')
    ax.grid()
    fig.savefig(join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + '_Pic_' + 'BackEMFPh1' + '.png'))
    data_array=[RotorPosition, BackEMFPh1,BackEMFLineToLine12]    
    io.savemat(join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + 'array' + 'BackEMFPh1' + '.mat',{"BackEMFPh1":data_array}))          
  # BEMF Export
  
  # Signal
    #List_Pos = []
    #List_TorqueVW_500rpm = []
    #for i in range(p_Torque_Pts+1):
    #    ex, Pos, Torque = mcApp.GetMagneticGraphPoint('TorqueVW', i)
    #    List_Pos.append(Pos)
    #    List_TorqueVW_500rpm.append(Torque)
    #if run_mode in ['OSL_run']:
    #    Pos_List_Len = len(List_Pos)
    #    o_Sig_Torque_Ripples = list_list_2_variant_signal([List_TorqueVW_500rpm], List_Pos)
    
# ### Current density
#     o_Current_Density = mcApp.GetVariable('ArmatureConductorCurrentDensity')
#     o_Current_Density = o_Current_Density[1]

### Re-set initial settings
    # mcApp.SetVariable("TorqueCalculation", False)      # Torque calculation
    # mcApp.SetVariable("AirgapMeshPoints_layers", 720)  # Number of mesh points in the airgap
    # mcApp.SetVariable("AirgapMeshPoints_mesh", 720)    # Number of mesh points at the airgap surface
    # mcApp.SetVariable("TorquePointsPerCycle", 30)      # Number of points to calculate for each cycle  
    # mcApp.SetVariable("MagneticSolver", 0)             # Reduced multi-static solver
           
### Mechanical context
    #mcApp.ShowMechanicalContext()
    #mcApp.DisplayScreen('Scripting')

### Mechanical: centrifugal stress calculation

  # Settings
    #mcApp.SetVariable("MechanicalOption_Magnets", 1)              # Magnets included
    #mcApp.SetVariable("MechanicalMeshLength_RotorLam", 0.25)      # Rotor lamination mesh size
    #mcApp.SetVariable("MechanicalMeshLength_Magnets", 0.5)        # Mesh size for magnets
    #mcApp.SetVariable("MechanicalMeshLength_RotorVoids", 0.2)     # Mesh size for flux barriers   
    #mcApp.SetVariable("Current_Shaft_Speed_RPM", 1.2*p_Speed_Max) # Rotational speed
    #mcApp.SetVariable("Shaft_Speed_[RPM]", 1.2*p_Speed_Max) 
     
  # Calculation
    # mcApp.DoMechanicalCalculation()
    # ex, Stress_Max  = mcApp.GetVariable('MaxStress_RotorLam')        # Peak stress
    # o_Stress_Safety = fun_Stress_Safety(p_Yield_Rotor, Stress_Max)   # Safety factor
   
### Close Motor-CAD (necessary when running designs in parallel)
    #mcApp.SaveToFile(mot_file_new_path)  # Save model
    #mcApp.Quit()                         # Close Motor-CAD
    #mcApp = 0                            # Reset mcApp variable  
    #time.sleep(0.5)                      # Freeze for 0.5s

### ----------------------------------------------      INITIALISATION (END)     ------------------------------------------------

### Responses to be drag and drop during 'OSL_setup' mode 
else:
    
  # Scalars
    # o_Cont_Torque_1krpm  = 0.
    # o_Cont_Torque_4krpm  = 0.
    o_Peak_Power_1krpm                           = 0.
    o_Peak_Power_4krpm                           = 0.
    o_Peak_Torque_1700rpm                        = 0.
    o_WLTP3_Eff                                  = 0.
 
    o_DutyCycleTotalEnergy_Electrical_Input     = 0.   
    o_DutyCycleTotalEnergy_Shaft_Output         = 0.  
    o_DutyCycleTotalEnergy_Electrical_Output    = 0. 
    o_DutyCycleTotalEnergy_Shaft_Input          = 0.      
        
    # o_Current_Density                            = 0.
    o_Torque_Density                             = 0.
    # o_Torque_Ripples                             = 0.
    o_Weight_Act                                 = 0.
    o_Weight_Mag                                 = 0.
    o_Weight_Rot_Core                            = 0.
    o_Weight_Stat_Core                           = 0.
    o_Weight_Wdg                                 = 0.
    # o_Stress_Safety                              = 0.
            
  # Signals
    o_Sig_Peak_Torque    = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
    o_Sig_Peak_Power     = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 
    # o_Sig_Torque_Ripples = list_list_2_variant_signal([[0]*(p_Torque_Pts+1)], [0]*(p_Torque_Pts+1))
