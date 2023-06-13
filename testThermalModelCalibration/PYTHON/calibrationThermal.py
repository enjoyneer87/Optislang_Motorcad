"""
# ----------------------------
# HDEV Calibration Thermal
# ----------------------------
# MOTOR-CAD : v15         
# OPTISLANG : v2023 R1 
# PYTHON    : v3.6.8
# ----------------------------
"""
# INITALIZATION
# Packages
# Running Mode & Working directories   
# USER-DEFINED  
    # Function
# Inputs 
# Calculation
#  Pre-calculations in PYTHON   
#  MOTOR-CAD  
#  SCREENSHOTS 
# INITIALISATION (END)  


### ----------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------             INITALIZATION          -----------------------------------------------
#-----------------------------------------------------      Packages        ----------------------------------------------------
import win32com.client
import os
import numpy as np
import time
from os import getcwd
from os.path import join, dirname, exists
from math import pi, sqrt
from scipy.io import loadmat
import matplotlib.pyplot as plt
from pyvariant import list_list_2_variant_signal

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
    OSL_DESIGN_NAME = 'Calibration0001'   
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


#-------------------------------------------------       Functions        ------------------------------------------------------

#-------------------------------------------------        Inputs        --------------------------------------------------------
### Motor-CAD options
Design_Name     = "HDEV_Model2Temp115"   # Reference Motor-CAD design
Visible_Opt     = 1.           # Set Motor-CAD visible
Message_Display = 2.           # Display all pop-up messages 
Save_Prompt     = 1.           # Never prompt to save file
CuboidalkValueDefinition=1
# MeasuredData = np.loadtxt('Z:/Thesis/Optislang_Motorcad/CalibrationBEMF/coMeasuredEMF.csv', delimiter=',')

### Post-processing
Pic_Export = 0      # Export geometry snapshots (0: No; 1: Yes)
### Input parameters for testing in IDE or initialisation in OSL Python node
if run_mode in ['OSL_setup', 'IDE_run']:

# Define Variable for motorcad
    
    K_Radial_User_A              = 0.8713
    K_Tangential_User_A          = 0.8713
    K_Axial_User_A               = 264.4
    K_Radial_User_F              = 0.8713
    K_Tangential_User_F          = 0.8713
    K_Axial_User_F               = 264.4
    K_Radial_User_R              = 0.8713
    K_Tangential_User_R          = 0.8713
    K_Axial_User_R               = 264.4

### Real solver run: 'IDE_run' mode or 'OSL_run' mode
if run_mode.endswith('run'):       
### --------------------------------------------           CALCULATIONS          -----------------------------------------------
### --------------------------------------------    Pre-calculations in PYTHON    ----------------------------------------------
### Geometry parameters
### ---------------------------------------------------      MOTOR-CAD     -----------------------------------------------------
### Load reference Motor-CAD file
    App = win32com.client.Dispatch("MotorCAD.AppAutomation")  # Launch Motor-CAD application
    App.Visible = Visible_Opt                                 # Set Motor-CAD visible or not
    App.SetVariable('MessageDisplayState', Message_Display)   # Set state of message display 
    App.SetVariable("SavePrompt", Save_Prompt)                # Remove autosave function or not
    myPID = os.getpid()											# Pass this process id to Motor-CAD so that Motor-CAD will 
    print(myPID)                                                # close when this process completes
    App.SetVariable("OwnerProcessID", myPID)
    App.ShowMessage(myPID)      
    mot_file_ref_path = join(refdir, Design_Name + ".mot")                       # Path to the reference *.mot file
    mot_file_new_path = join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + ".mot") # Path to the new *.mot file
    print("[INFO] Load reference MCAD file from: ", mot_file_ref_path)                                    
    App.LoadFromFile(mot_file_ref_path)                                        # Load reference Motor-CAD file
    print("[INFO] Design file saved as: ", mot_file_new_path)                                    
    App.SaveToFile(mot_file_new_path)                                          # Save in new location
    mot_file_dir = join(wdir, Design_Name + '_' + OSL_DESIGN_NAME)

### Thermal context    
    App.ShowThermalContext()
    
    App.SetVariable('CuboidalkValueDefinition'    ,    CuboidalkValueDefinition   )
    App.SetVariable('K_Radial_User_A'        , K_Radial_User_A        )         
    App.SetVariable('K_Tangential_User_A'    , K_Tangential_User_A    )           
    App.SetVariable('K_Axial_User_A'         , K_Axial_User_A         )            
    App.SetVariable('K_Radial_User_F'        , K_Radial_User_F        ) 
    App.SetVariable('K_Tangential_User_F'    , K_Tangential_User_F    )             
    App.SetVariable('K_Axial_User_F'         , K_Axial_User_F         )             
    App.SetVariable('K_Radial_User_R'        , K_Radial_User_R        )             
    App.SetVariable('K_Tangential_User_R'    , K_Tangential_User_R    )             
    App.SetVariable('K_Axial_User_R'         , K_Axial_User_R         )    
       
### Check the geometry is valid
    success = App.CheckIfGeometryIsValid(0)
    if success == 0:
        x_axis = [0.0, 20.0, 45.0]
        y_axes = [101, 95, 90]
        # If not valid, generate zero outputs instead of getting an error message in optiSLang
        if run_mode in ['OSL_run']:
                    o_winding_temp_average_transient =list_list_2_variant_signal([y_axes], x_axis)
        App.SaveToFile(mot_file_new_path)  # Save design
        App.Quit()                         # Close Motor-CAD
        App = 0                            # Reset App variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: geometry not valid'.format(OSL_DESIGN_NAME))
    

### Export snapshots
    if Pic_Export:
        for screenname in ['Radial', 'Axial', 'StatorWinding']:
            App.SaveScreenToFile(screenname, join(wdir, Design_Name + '_' + OSL_DESIGN_NAME + '_Pic_' + screenname + '.png'))
    App.DisplayScreen('Scripting')   

### Extract active weights
    App.DoWeightCalculation()                                            # Weight calculation
    ex, o_Weight_Mag       = App.GetVariable("Weight_Calc_Magnet")       # Magnet's mass 
    ex, o_Weight_Wdg       = App.GetVariable("Weight_Calc_Copper_Total") # Winding's mass 
    ex, o_Weight_Stat_Core = App.GetVariable("Weight_Calc_Stator_Lam")   # Stator core's mass  
    ex, o_Weight_Rot_Core  = App.GetVariable("Weight_Calc_Rotor_Lam")    # Rotor core's mass
    ex, Weight_Shaft       = App.GetVariable("Weight_Shaft_Active")      # Shaft's mass  
    ex, Weight_Act         = App.GetVariable("Weight_Calc_Total")        # Active mass
    o_Weight_Act           = Weight_Act - Weight_Shaft                     # Shaft's mass retrieved
    
   
### Save design
    App.SaveToFile(mot_file_new_path)
    
### %% Duty Cycle Setting 

#  CreateDutyCycleData
    mcadTimePeriod = [5, 10, 10, 10, 10]
    Npoints = [4] * 5
    Speed_start = [1700, 1000, 500, 0, 0]
    Speed_End = [1000, 500, 0, 0, 0]
    torqueStart= [0] *5
    torqueEnd = [0] *5

    DutyCycleStr = {
    'Duty_Cycle_Points': Npoints,
    'Duty_Cycle_Num_Periods': len(mcadTimePeriod),
    'Duty_Cycle_Speed_Start': Speed_start,
    'Duty_Cycle_Speed_End': Speed_End,
    'Duty_Cycle_Time': mcadTimePeriod,
    'Duty_Cycle_Torque_Start' : torqueStart,
    'Duty_Cycle_Torque_End' : torqueEnd,    
    }


    ## SetDutyCycleData

    App.SetVariable('Duty_Cycle_Num_Periods',DutyCycleStr['Duty_Cycle_Num_Periods'])

    for array_index in range(len(DutyCycleStr['Duty_Cycle_Points'])):
        
        App.SetArrayVariable('Duty_Cycle_Time', array_index, DutyCycleStr['Duty_Cycle_Time'][array_index])
        App.SetArrayVariable('Duty_Cycle_Points', array_index, DutyCycleStr['Duty_Cycle_Points'][array_index])
        App.SetArrayVariable('Duty_Cycle_Speed_Start', array_index, DutyCycleStr['Duty_Cycle_Speed_Start'][array_index])
        App.SetArrayVariable('Duty_Cycle_Speed_End', array_index, DutyCycleStr['Duty_Cycle_Speed_End'][array_index])
        App.SetArrayVariable('Duty_Cycle_Torque_Start', array_index, DutyCycleStr['Duty_Cycle_Torque_Start'][array_index])
        App.SetArrayVariable('Duty_Cycle_Torque_End', array_index, DutyCycleStr['Duty_Cycle_Torque_End'][array_index])    

    ### Save design    
    App.SaveToFile(mot_file_new_path)    
        
    ## [Thermal] DoTransientAnalysis
    App.DoTransientAnalysis()
        # Raise exception if wrong performance data
        # if o_WLTP3_Eff < 0:
        #         App.SaveToFile(mot_file_new_path)  # Save design   
        #     App.Quit()                         # Close Motor-CAD
        #     App = 0                            # Reset App variable  
        #     time.sleep(0.5)                      # Frozen for 0.5s
        #     raise Exception('[ERROR] {}: DoTransientAnalysis failed'.format(OSL_DESIGN_NAME))
    
                
  # Signal
    
    setGraphName = ['Wdg (Average) (C11)']
    dataIndex = 0
    NumGraphPoints = 4 * 5

    List_xvalue = []
    List_valueforGraph = []
    
    for loop in range(NumGraphPoints):
        success, x, y = App.GetTemperatureGraphPoint(setGraphName[dataIndex], loop)
        List_xvalue.append(x)
        List_valueforGraph.append(y)
        x_axis = List_xvalue
        y_axes = List_valueforGraph
    if run_mode in ['OSL_run']:
        timeListLen = len(x_axis)
        o_winding_temp_average_transient =list_list_2_variant_signal([y_axes], x_axis)


    # plt.plot(MeasuredData)
    # plt.plot(List_PhaseEMFwo3rd)
    # plt.savefig('PhaseEMFwaveform.png')    

    # plt.figure()
    # plt.bar(range(len(pos_fft_data1)), pos_fft_data1, alpha=0.5, label='Measured')
    # plt.bar(range(len(data2_fftposi)), data2_fftposi, alpha=0.5, label='MotorCAD')
    # n = len(data2_fftposi)
    # plt.xlabel('order')
    # plt.legend()
    # xticklabels=np.arange(0, (n+1)//2,3).tolist()
    # xticks=np.arange(0,len(data2_fftposi),3)
    # plt.xticks(xticks,xticklabels)
    # plt.savefig('ComparePhaseFFT.png')    
    
### Close Motor-CAD (necessary when running designs in parallel)
    App.SaveToFile(mot_file_new_path)  # Save model
    App.Quit()                         # Close Motor-CAD
    App = 0                            # Reset App variable  
    time.sleep(0.5)                      # Freeze for 0.5s

### ----------------------------------------------      INITIALISATION (END)     ------------------------------------------------



else:    

    x_axis = [0.0, 20.0, 45.0]
    y_axes = [101, 95, 90]
    o_winding_temp_average_transient =list_list_2_variant_signal([y_axes], x_axis)
    # SIGNAL [3, 2]
    # (1,0) - (0.1,0) (1.3,0)
    # (2,0) - (0.2,0) (1.2,0)
    # (3,0) - (0.3,0) (1.1,0)
