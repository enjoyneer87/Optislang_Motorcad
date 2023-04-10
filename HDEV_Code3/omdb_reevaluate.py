# -*- coding: utf-8 -*-

# This Python example builds a optiSLang workflow to run a revaluation with an 
# existing optiSLang omdb file, e.g. to extract additional signal information.

# Run this script in optiSLang's Python Console. 
# It automatically asks for the existing omdb file, builds the workflow and applies necessary settings.
# Finally you just need to add you additional signal extractions, criteria, ....  

from os.path import basename
from py_omdb import PyOMDB
from py_os_parameter import PyParameterManager
from py_os_design import PyOSDesignContainer, PyOSDesignPoint


# get omdb file
omdb_file = gui.get_open_file_name('Please select optiSLang Postprocessing file', '', 'Monitoring database (*.omdb)' , '', gui.FileDialogOption.DONTCONFIRMOVERWRITE)
if not len(omdb_file):
    raise Exception('no file')

# read file content
file_name_omdb = basename(omdb_file)    
omdb = PyOMDB(omdb_file)
pm = omdb.parameter_manager
dc = omdb.design_container

# setup reevaluation system
sensitivity_system = actors.SensitivityActor('Reevaluate {}'.format(file_name_omdb))
sensitivity_system.auto_save_mode = AS_ACTOR_FINISHED
sensitivity_system.parameter_manager = pm
sensitivity_system.start_designs = dc
sensitivity_system.solve_duplicated = True
sensitivity_system.solve_start_designs_again = True
sensitivity_system.preserve_start_design_ids = True
sensitivity_system.dynamic_sampling = False

# setup omdb node
omdb_node = actors.CustomIntegrationActor('optislang_omdb')
omdb_node.name = file_name_omdb
omdb_node.path = ProvidedPath(omdb_file)

# get parameter and response names
for d in dc:
    for n, v in d.get_parameters():
        dp = PyOSDesignPoint()
        dp.add('Name', n)
        dp.add('Value', v)
        omdb_node.add_parameter((n, v), actors.CustomizedBaseInfo(dp, actors.IntegrationDirection.DIRECTION_INPUT))
    for n, v in d.get_responses():
        dp = PyOSDesignPoint()
        dp.add('Name', n)
        dp.add('Value', v)
        omdb_node.add_response((n, v), actors.CustomizedBaseInfo(dp, actors.IntegrationDirection.DIRECTION_OUTPUT))
    break

# build workflow    
add_actor(sensitivity_system)
sensitivity_system.add_actor(omdb_node)
connect(sensitivity_system, 'IODesign', omdb_node, 'IDesign')
connect(omdb_node, 'ODesign', sensitivity_system, 'IIDesign')
