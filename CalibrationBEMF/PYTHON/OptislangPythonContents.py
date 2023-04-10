#-------------------------------------------------        Inputs        --------------------------------------------------------
### Motor-CAD options
### Materials
### Post-processing



### --------------------------------------------           CALCULATIONS          -----------------------------------------------
### --------------------------------------------    Pre-calculations in PYTHON    ----------------------------------------------
### Geometry parameters
### ---------------------------------------------------      MOTOR-CAD     -----------------------------------------------------
### Load reference Motor-CAD file
### EMag context

### Change to ratio mode to edit the geometry
### Assign geometry parameters
### Check the geometry is valid
    
### Assign winding parameters
### Assign initial calculation settings

### Extract active weights
### Save design
### EMag context
### EMF calculation 
### Re-set initial settings

### Re-set initial settings
    # mcApp.SetVariable("TorqueCalculation", False)      # Torque calculation
    # mcApp.SetVariable("AirgapMeshPoints_layers", 720)  # Number of mesh points in the airgap
    # mcApp.SetVariable("AirgapMeshPoints_mesh", 720)    # Number of mesh points at the airgap surface
    # mcApp.SetVariable("TorquePointsPerCycle", 30)      # Number of points to calculate for each cycle  
    # mcApp.SetVariable("MagneticSolver", 0)             # Reduced multi-static solver
           

### Mechanical context
### Mechanical: centrifugal stress calculation
### ----------------------------------------------      INITIALISATION (END)     ------------------------------------------------
# else:
    
  # Scalars
    o_Cont_Torque_1krpm  = 0.

  # Signals
    # o_Sig_Peak_Torque    = list_list_2_variant_signal([[0]*Speed_Lab_Len], Speed_Lab) 


### Shows automatically after assigning options for the saturation & loss models)
### Lab: model Build tab 
### Lab: peak performance
  # Settings
  # Calculation & Data management
  # Extract specific values
  # Raise exception if wrong performance data
  # Key performance indicators
  # Signals to be read in OSL 
### Lab: efficiency over WLTP-3 drive cycle 
  # Settings
  # Calculation & Post processing
  # Raise exception if wrong performance data

### Lab: continuous performance
  # Settings
  # Calculation & Post processing
  # Raise exception if negative value    