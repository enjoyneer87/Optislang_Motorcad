"""
# ----------------------------
# HDEV Calibration BEMF
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
import numpy as np

def zeroOut3ncoeffs(signal):
    # ?? ???? ?? ??
    n = len(signal)
    # ?? ???? FFT ??
    signal_fft = np.fft.fft(signal)
    # 0?? ???? ?? ?? fftshift ??? ?????.
    signal_fft = np.fft.fftshift(signal_fft)
    # 3? ??? ???? ?? ??
    for j in range(n//12):
        signal_fft[n//2-3*(j+1)]=0
        signal_fft[n//2+3*(j+1)]=0
    # inverse FFT? ???? filtering? ??? ??
    filtered_signal = np.fft.ifft(np.fft.ifftshift(signal_fft))
    return filtered_signal

def compare_fft(data1, data2):
    # FFT ??
    fft_data1 = np.fft.fft(data1)
    fft_data2 = np.fft.fft(data2)

    # FFT ?? ??
    abs_fft_data1 = 2.0 / len(data1) * np.abs(fft_data1)
    abs_fft_data2 = 2.0 / len(data2) * np.abs(fft_data2)

    # ??? ?? ??? ?? ??
    diff = abs_fft_data1 - abs_fft_data2

    # ??? ???? ???? ??? ??
    error = np.sqrt(np.sum(np.power(diff, 2)))

    return diff,error

def euclidean_norm(signal1, signal2):
    # ?? ?? ??
    diff = signal1 - signal2
    square_diff = np.power(diff, 2)
    
    # ?? ??? ? ??
    sum_square_diff = np.sum(square_diff)
    
    # ?? ??? ?? ??? ??
    norm = np.sqrt(sum_square_diff)
    
    return norm

def rmse(y_true, y_pred):
    # ???? ???? ?? ??
    diff = y_true - y_pred
    # ?? ?? ??
    square_diff = np.power(diff, 2)
    # ?? ??? ?? ??
    mean_square_diff = np.mean(square_diff)
    # RMSE ??
    rmse = np.sqrt(mean_square_diff)
    return rmse

def min_max_normalize(data):
    # ???? ???? ??? ??
    min_val = np.min(data)
    max_val = np.max(data)
    # ????? ? ?, ????? ???? ???? 0~1 ??? ??? ???
    norm_data = (data - min_val) / (max_val - min_val)
    return norm_data


#-------------------------------------------------        Inputs        --------------------------------------------------------
### Motor-CAD options
Design_Name     = "HDEV_Model1"   # Reference Motor-CAD design
Visible_Opt     = 1.           # Set Motor-CAD visible
Message_Display = 2.           # Display all pop-up messages 
Save_Prompt     = 1.           # Never prompt to save file

MeasuredData = np.loadtxt('Z:/Thesis/Optislang_Motorcad/CalibrationBEMF/coMeasuredEMF.csv', delimiter=',')

### Post-processing
Pic_Export = 0      # Export geometry snapshots (0: No; 1: Yes)
### Input parameters for testing in IDE or initialisation in OSL Python node
if run_mode in ['OSL_setup', 'IDE_run']:
    p_EMF_Pts=120
# Magnet Temperature
    i_Magnet_Temperature = 20    
#% Stacklength Coefficient
    i_Stacking_Factor_Stator         =0.97    
    i_Stacking_Factor_Rotor          =0.97
#  Stackingfactorcalculation  StackingFactor_Magnetics
    # i_StackingFactor_Magnetics              =2      # 0 ignore, 1: Axial length, 2:saturation(default) 
# Manufacturing Factors:  
    i_ArmatureEWdgMLT_Multiplier                               =1
    i_ArmatureEWdglnductance_Multiplier                        =1
    i_Magnet_Br_Multiplier                                     =0.95     
    # i_ArmatureEWdgMLT_Aux_Multiplier                           =1
    # i_ArmatureEWdglnductance_Aux_Multiplier                    =0.97    
# Length Adjustment Factors
    StatorSaturationMultiplier                               =0.95 
    RotorSaturationMultiplier                                =0.95
    MagneticAxialLengthMultiplier                            =1   
    #MagneticAxialLength_Array  
    #NumberStatorLamination
    #NumberRotorLamination
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
### EMag context
    App.ShowMagneticContext()          
    App.DisplayScreen('Scripting')     # Switch to a tab where no parameter is adjusted 
### Assign initial calculation settings
    App.SetVariable("BackEMFCalculation", True)               # OC calculations deactivated


### EMF calculation 
  # Settings      
    # App.SetVariable("BackEMFPointsPerCycle", p_EMF_Pts) # Number of points to be used for each cycle  
    # App.SetVariable("BackEMFNumberCycles", p_EMF_NumberCycles) # Number of points to be used for each cycle  

    App.SetVariable("Stacking_Factor_[Stator]",i_Stacking_Factor_Stator)
    App.SetVariable("Stacking_Factor_[Rotor]",i_Stacking_Factor_Rotor)
    # App.SetVariable("StackingFactor_Magnetics",i_StackingFactor_Magnetics)
    # App.SetVariable("ArmatureEWdgMLT_Multiplier",i_ArmatureEWdgMLT_Multiplier)
    # App.SetVariable("ArmatureEWdglnductance_Multiplier",i_ArmatureEWdglnductance_Multiplier)
    App.SetVariable("Magnet_Br_Multiplier",i_Magnet_Br_Multiplier)
    App.SetVariable("StatorSaturationMultiplier",StatorSaturationMultiplier)
    App.SetVariable("RotorSaturationMultiplier",RotorSaturationMultiplier)
    App.SetVariable("MagneticAxialLengthMultiplier",MagneticAxialLengthMultiplier)
    App.SetVariable("Magnet_Temperature",i_Magnet_Temperature)  
    o_normalDataError=0
    o_error          = 0.
    normalRmse      =  0.
### Save design    
    App.SaveToFile(mot_file_new_path)    

    # Calculation & Post processing
    App.DoMagneticCalculation()   
    App.SaveResults('Emagnetic')                  
  # Signal
    success,p_EMF_Pts=App.GetVariable('BackEMFPointsPerCycle')
    List_Pos = []
    List_PhaseEMF = []
    for i in range(p_EMF_Pts+1):
        ex, Pos, PhaseEMF = App.GetMagneticGraphPoint('BackEMFPh1', i)
        List_Pos.append(Pos)
        List_PhaseEMF.append(PhaseEMF)
    
    
    List_Pos=List_Pos[0:p_EMF_Pts]
    List_PhaseEMF=List_PhaseEMF[0:p_EMF_Pts]
    List_PhaseEMFwo3rd=zeroOut3ncoeffs(List_PhaseEMF)
    List_PhaseEMFwo3rd = List_PhaseEMFwo3rd.tolist()
    

    
    
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

    if run_mode in ['OSL_run']:
        normalEMFsignal1 = min_max_normalize(MeasuredData)
        normalEMFsignal2 = min_max_normalize(List_PhaseEMFwo3rd)
        diff_normalDataFFT,o_normalDataError=compare_fft(normalEMFsignal1,normalEMFsignal2)
        diff_FFT,o_error=compare_fft(MeasuredData,List_PhaseEMFwo3rd)
        Len_Pos_List = len(List_Pos)
        Len_FFT_List = len(diff_FFT)
   
        o_Sig_BackEMF = list_list_2_variant_signal([List_PhaseEMF], List_Pos)
        o_Sig_BackEMF_wo3rd = list_list_2_variant_signal([List_PhaseEMFwo3rd], List_Pos)

        normalRmse=rmse(normalEMFsignal1, normalEMFsignal2)

        # o_data2_fftposi = list_list_2_variant_signal([data2_fftposi], List_Pos)
      # Raise exception if wrong performance data
    if o_error < 0:
        mcApp.SaveToFile(mot_file_new_path)  # Save design   
        mcApp.Quit()                         # Close Motor-CAD
        mcApp = 0                            # Reset mcApp variable  
        time.sleep(0.5)                      # Frozen for 0.5s
        raise Exception('[ERROR] {}: error calculation error'.format(OSL_DESIGN_NAME))
    
### Close Motor-CAD (necessary when running designs in parallel)
    App.SaveToFile(mot_file_new_path)  # Save model
    App.Quit()                         # Close Motor-CAD
    App = 0                            # Reset App variable  
    time.sleep(0.5)                      # Freeze for 0.5s

### ----------------------------------------------      INITIALISATION (END)     ------------------------------------------------

### Responses to be drag and drop during 'OSL_setup' mode 
else:
    
  # Scalars
    o_error  = 0.
    normalRmse = 0.
    o_normalDataError=0
  # Signals

    o_Sig_BackEMF = list_list_2_variant_signal([[0]*(p_EMF_Pts)], [0]*(p_EMF_Pts))
    o_Sig_BackEMF_wo3rd = list_list_2_variant_signal([[0]*(p_EMF_Pts)], [0]*(p_EMF_Pts))
