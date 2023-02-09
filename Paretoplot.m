

sensitivityTable=readtable('Samoo_HDEV_low_fidelity_sensitivity.csv');
lowfidel_designtable=readtable('Samoo_HDEV_low_fidelity_design_table.csv');
lowfidel_designtable_true=readtable('Samoo_HDEV_low_fidelity_design_table_true.csv');
lowfidel_designtable_pareto=readtable('Samoo_HDEV_low_fidelity_pareto.csv');
grayColor = [.7 .7 .7];
paretofigure=figure(1)
scatter(lowfidel_designtable.obj_o_Weight_Act,lowfidel_designtable.obj_o_Wh_Loss);
aAxes=paretofigure.Children;

sc=aAxes.Children;
sc.DisplayName='NonFeasible';
sc.MarkerEdgeColor=grayColor;
hold on
legend
feasiblePlot=scatter(lowfidel_designtable_true.obj_o_Weight_Act,lowfidel_designtable_true.obj_o_Wh_Loss);
feasiblePlot.MarkerFaceColor='b';
feasiblePlot.DisplayName='Feasible';
hold on
lowfidel_designtable_paretoPlot=plot(lowfidel_designtable_pareto.obj_o_Weight_Act,lowfidel_designtable_pareto.obj_o_Wh_Loss);
lowfidel_designtable_paretoPlot.Color='r';
lowfidel_designtable_paretoPlot.Marker='*';
lowfidel_designtable_paretoPlot.DisplayName='Pareto'
sensTablePlot=scatter(sensitivityTable.obj_o_Weight_Act,sensitivityTable.obj_o_Wh_Loss);
sensTablePlot.DisplayName='DOE';

set(gca,'FontName','Times New Roman','FontSize',12)
grid on
legend
ax=gca;
ax.YLabel.String="Driving Cycle Energy Consumtions [Wh]";
ax.XLabel.String="Active Part Weight [kg]";


