#my prj
sed -i 's@:KISYS3DMOD_MAS:@${MAS_3DMODEL_DIR}/@g' *.kicad_pcb
sed -i 's@MAS_IC_PowerSupply:DKA30B-15@MAS_Module:DKA30B-15@g' *.kicad_sch

#prj
sed -i 's@${KISYS3DMOD}@${KICAD6_3DMODEL_DIR}@g' *.kicad_pcb
sed -i 's@Device:CP_Small@Device:C_Polarized_Small@g' *.kicad_sch