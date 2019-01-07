Macro "MC Central"

Switch = 1 // 1- household trips, 2 - university trips

// NOTE to RSG - assuming these paths will be passed into the final version of the macro from the main model
DCpath = "c:\\projects\\washtenaw model\\20181130\\20181130\\05-hb-modechc\\04-hb-destchc\\" // FLAG hardcoded path
SKpath = "c:\\projects\\washtenaw model\\wats_model-master (2)\\wats_model-master\\2-scenarios\\working\\02-network\\"// FLAG hardcoded path
MCpath = "c:\\projects\\washtenaw model\\Mode Choice\\" // FLAG hardcoded path
ParamPath = "C:\\Projects\\Washtenaw Model\\wats_model-master (2)\\wats_model-master\\1-Model-System-Files\\5-Parameters\\" // FLAG hardcoded path

// Auto Mode Costs - NOTE, we should put these cost assumptions in the user interface for easier updating
CostPerMI_Auto = 0.55
CostPerTT_ZipCar = 0.142
CostPerMI_Uber = 0.97
CostPerTT_Uber = 0.14
CostInit_Uber = 1.98
CostArray = {CostPerMI_Auto, CostPerTT_ZipCar, CostPerMI_Uber, CostPerTT_Uber, CostInit_Uber}

RunMacro("Add MC Skim Cores", SKpath, MCpath, CostArray)

if Switch =1 then do 
    RunMacro("Add HHMC Cores", DCpath)
    Seg = {"AS2","AS1","AS0"}
    MinPerUSD = {3.68,5.80,8.73}
    AS2cores ={"HBW_AGW_PK", "HBW_AGW_OP", "WHBO_AGW_PK", "WHBO_AGW_OP", "HBOLOT_AGW_PK", "HBOLOT_AGW_OP", "HBOSOT_AGW_PK", "HBOSOT_AGW_OP"}
    AS1cores ={"HBW_ALW_PK", "HBW_ALW_OP", "WHBO_ALW_PK", "WHBO_ALW_OP", "HBOLOT_ALW_PK", "HBOLOT_ALW_OP", "HBOSOT_ALW_PK", "HBOSOT_ALW_OP"}
    AS0cores ={"HBW_ZA_PK", "HBW_ZA_OP", "WHBO_ZA_PK", "WHBO_ZA_OP", "HBOLOT_ZA_PK", "HBOLOT_ZA_OP", "HBOSOT_ZA_PK", "HBOSOT_ZA_OP"}
    MCSegments = {AS2cores, AS1cores, AS0cores}
    DCfiles ={"PA_TripsAS2.mtx","PA_TripsAS1.mtx","PA_TripsAS0.mtx"}
    models = {"AS2_MC_NLM.mdl","AS1_MC_NLM.mdl","AS0_MC_NLM.mdl"}

    for i = 1 to Arraylength(Seg) do
        RunMacro("Cost to Minutes" ,MCpath, MinPerUSD[i])   
        RunMacro("Apply MC", MCpath, DCpath, models[i], MCSegments[i], DCfiles[i],Seg[i], ParamPath)
    end
    RunMacro("FinalizeHHMC", MCpath)
end

if Switch = 2 then do 
    RunMacro("Add UMC Cores", DCpath)
    Seg = {"AS1","AS0"}
    MinPerUSD = {3.82,8.73}
    UAS1cores ={"U_CBC_AS1_PK", "U_CBC_AS1_OP", "U_HBC_AS1_PK", "U_HBC_AS1_OP", "U_CBO_AS1_PK", "U_CBO_AS1_OP", "U_HBO_AS1_PK", "U_HBO_AS1_OP"}
    UAS0cores ={"U_CBC_AS0_PK", "U_CBC_AS0_OP", "U_HBC_AS0_PK", "U_HBC_AS0_OP", "U_CBO_AS0_PK", "U_CBO_AS0_OP", "U_HBO_AS0_PK", "U_HBO_AS0_OP"}
    MCSegments = {UAS1cores, UAS0cores}
    DCfiles ={"PA_TripsAS1.mtx","PA_TripsAS0.mtx"}
    models = {"AS1_UMC_NLM.mdl","AS0_UMC_NLM.mdl"}

    for i = 1 to Arraylength(Seg) do
        RunMacro("Cost to Minutes" ,MCpath, MinPerUSD[i])   
        RunMacro("Apply MC", MCpath, DCpath, models[i], MCSegments[i], DCfiles[i],Seg[i], ParamPath)
    end
        RunMacro("FinalizeUMC", MCpath)
end

endMacro

//============================================================================================
//  Add MC Skim Cores - Combine Bus and Auto skim info into a single matrix, compute costs
//============================================================================================
Macro "Add MC Skim Cores" (input_path, mcfolder, UnitCosts)
    RunMacro("TCB Init")

    mb = OpenMatrix(input_path + "impAM_WLKTRN.mtx", )
    mc = CreateMatrixCurrency(mb, "Fare", "RCIndex", "RCIndex", )
    new_mat = CopyMatrix(mc, {{"File Name",mcfolder + "MC_Combined_Skims.mtx"},{"Label", "MC Inputs"},{"Indices", "All"}})

    addedcores = {"BUS_AVAIL", "BUS_OVTT", "BUSCOSTmin", "CARSHARE_AVAIL", "DIST", "WalkTT", "BikeTT", "ZipTT", "ZipCOST", "ZipOVTT", "ZipCOSTmin", "HHautoTT", "HHautoCOST", "HHautoOVTT", "HHautoCOSTmin",
                "srHHautoTT", "srHHautoCOST", "srHHautoCOSTmin", "srHHautoOVTT", "CarpoolTT", "CarpoolCOST","CarpoolCOSTmin", "CarpoolOVTT", "UberTT","UberCOST", "UberCOSTmin", "UberOVTT"}
    
    mb2 = OpenMatrix(mcfolder + "MC_Combined_Skims.mtx", )

    for j=1 to ArrayLength(addedcores) do
        AddMatrixCore(mb2, addedcores[j])
    end

    mcs = OpenMatrix(input_path + "CarShareAvailable.mtx", )
    mc3 = CreateMatrixCurrency(mcs, "CarShareAvail", "RCIndex", "RCIndex", )

    mc_arrayB = CreateMatrixCurrencies(mb2, "RCIndex", "RCIndex", )
    mc_arrayB.BUS_OVTT := mc_arrayB.[Initial Wait Time] + mc_arrayB.[Transfer Wait Time] + mc_arrayB.[Transfer Walk Time] + mc_arrayB.[Access Walk Time] + mc_arrayB.[Egress Walk Time]
    mc_arrayB.BUS_AVAIL := min(mc_arrayB.Fare,1)
    mc_arrayB.CARSHARE_AVAIL := mc3

    CoreNames = GetMatrixCoreNames(mb2)

    ma = OpenMatrix(input_path + "impdaAM.mtx", )
    mc_arrayA = CreateMatrixCurrencies(ma, "Origin", "Destination", )

    // Travel times by mode
    mc_arrayB.DIST := mc_arrayA.[Length (Skim)]
    mc_arrayB.WalkTT := (mc_arrayB.DIST/3)*60
    mc_arrayB.BikeTT := (mc_arrayB.DIST/12)*60
    mc_arrayB.ZipTT := mc_arrayA.[AM_Time (Skim)]
    mc_arrayB.HHautoTT := mc_arrayA.[AM_Time (Skim)]
    mc_arrayB.srHHautoTT := mc_arrayA.[AM_Time (Skim)] 
    mc_arrayB.CarpoolTT := mc_arrayA.[AM_Time (Skim)]  
    mc_arrayB.UberTT := mc_arrayA.[AM_Time (Skim)]

    // OVTT by mode
    mc_arrayB.ZipOVTT := 15 // includes walk access time and vehicle aquisition
    mc_arrayB.HHautoOVTT := 2 
    mc_arrayB.srHHautoOVTT := 10
    mc_arrayB.CarpoolOVTT := 15
    mc_arrayB.UberOVTT := 7 // Wait time for pickup, With future information this should vary by TAZ

    // Cost by mode
    mc_arrayB.HHautoCOST := mc_arrayB.DIST * UnitCosts[1]
    mc_arrayB.ZipCOST := mc_arrayB.ZipTT * UnitCosts[2] 
    mc_arrayB.srHHautoCOST := mc_arrayB.DIST * UnitCosts[1]
    mc_arrayB.CarpoolCOST := mc_arrayB.DIST * UnitCosts[1]
    mc_arrayB.UberCOST := mc_arrayB.DIST * UnitCosts[3] + mc_arrayB.UberTT * UnitCosts[4]  + UnitCosts[5]

endMacro

//============================================================================================
//  Add HHMC Cores - Put houehold DC outputs into the format needed for MC
//============================================================================================
Macro "Add HHMC Cores" (input_path)
    RunMacro("TCB Init")
    
    Seg = {"AS2","AS1","AS0"}
    PAfiles = {"PA_TripsAS2.mtx","PA_TripsAS1.mtx","PA_TripsAS0.mtx"}
    newcores = {"HBW_PK", "HBW_OP", "HBOWT_PK","HBOWT_OP", "HBOSOT_PK", "HBOSOT_OP", "HBOLOT_PK", "HBOLOT_OP"}
    DCFiles = {"PA_HBW_PK.mtx","PA_HBW_OP.mtx","PA_HBOWT_PK.mtx","PA_HBOWT_OP.mtx","PA_HBOSOT_PK.mtx","PA_HBOSOT_OP.mtx","PA_HBOLOT_PK.mtx","PA_HBOLOT_OP.mtx"}

for h=1 to Arraylength(Seg) do
    for i =1 to Arraylength(newcores) do
    // Add Matrix Core
        Opts = null
        Opts.Input.[Input Matrix] = input_path + PAfiles[h]
        Opts.Input.[New Core] = newcores[i]
        ok = RunMacro("TCB Run Operation", "Add Matrix Core", Opts, &Ret)

    // Fill from model output
        Opts = null
        Opts.Input.[Target Currency] = {input_path + PAfiles[h], newcores[i], "Row ID's", "Col ID's"}
        Opts.Input.[Source Currencies] = {{input_path + DCFiles[i], Seg[h],  "TAZID", "TAZID"}}
        Opts.Global.[Missing Option].[Force Missing] = "No"
        ok = RunMacro("TCB Run Operation", "Merge Matrices", Opts, &Ret)

    end
end

endMacro

//============================================================================================
//  Add University MC Cores - Put University DC outputs into the format needed for MC
//============================================================================================
Macro "Add UMC Cores" (input_path)
    RunMacro("TCB Init")
    
    Seg = {"AS1","AS0"}
    PAfiles = {"PA_TripsAS1.mtx","PA_TripsAS0.mtx"}
    newcores = {"U_CBC_PK", "U_CBC_OP", "U_HBC_PK","U_HBC_OP", "U_CBO_PK", "U_CBO_OP", "U_HBO_PK", "U_HBO_OP"}
    DCFiles = {"PA_PK_U_CBC.mtx","PA_OP_U_CBC.mtx","PA_PK_U_HBC.mtx","PA_OP_U_HBC.mtx","PA_PK_U_CBO.mtx","PA_OP_U_CBO.mtx","PA_PK_U_HBO.mtx","PA_OP_U_HBO.mtx"}

for h=1 to Arraylength(Seg) do
    for i =1 to Arraylength(newcores) do
    // Add Matrix Core
        Opts = null
        Opts.Input.[Input Matrix] = input_path + PAfiles[h]
        Opts.Input.[New Core] = newcores[i]
        ok = RunMacro("TCB Run Operation", "Add Matrix Core", Opts, &Ret)
        
    // Fill from model output
        Opts = null
        Opts.Input.[Target Currency] = {input_path + PAfiles[h], newcores[i], "Row ID's", "Col ID's"}
        Opts.Input.[Source Currencies] = {{input_path + DCFiles[i], Seg[h],  "Rows", "Columns"}}
        Opts.Global.[Missing Option].[Force Missing] = "No"
        ok = RunMacro("TCB Run Operation", "Merge Matrices", Opts, &Ret)
        
    end
end


endMacro

//============================================================================================
//  Const to Minutes - Compute costs as equivalent minutes for use in utility equation
//============================================================================================
Macro "Cost to Minutes" (MCpath, cost)

    mc = OpenMatrix(MCpath + "MC_Combined_Skims.mtx", )
    mc_array = CreateMatrixCurrencies(mc, "RCIndex", "RCIndex", )
    mc_array.ZipCOSTmin := mc_array.ZipCOST * cost
    mc_array.HHautoCOSTmin := mc_array.HHautoCOST * cost
    mc_array.srHHautoCOSTmin := mc_array.srHHautoCOST * cost
    mc_array.CarpoolCOSTmin := mc_array.CarpoolCOST * cost
    mc_array.UberCOSTmin := mc_array.UberCOST * cost
    mc_array.BUSCOSTmin := 1.5 * mc_array.Fare * cost

endMacro

//============================================================================================
//  Apply MC - Applies the relevant nested mode choice model
//============================================================================================
Macro "Apply MC" (input_path, input_path2, model, segments, DCfile, seg, parampath)
    RunMacro("TCB Init")
    output_path = input_path
    P_out = "Probabilities.mtx"
    A_out = "Applied_Totals.MTX"

// STEP 1: NestedLogitEngine
    Opts = null
    Opts.Global.[Missing Method] = "Drop Mode"
    Opts.Global.[Base Method] = "On View"
    Opts.Global.[Small Volume To Skip] = 0.001
    Opts.Global.[Utility Scaling] = "By Parent Theta"
    Opts.Global.ShadowIterations = 10
    Opts.Global.ShadowTolerance = 0.001
    Opts.Global.Model = parampath + model
    Opts.Global.Segments = segments
    Opts.Flag.ShadowPricing = 0
    Opts.Flag.Aggregate = 1
    Opts.Input.[Matrix2 Matrix] = input_path + "MC_Combined_Skims.mtx"
    Opts.Input.[Matrix1 Matrix] = input_path2 + DCfile
    Opts.Output.[Probability Matrices] = {{{"Label", segments[1] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[1]+"_" +seg+ "_HH_Probabilities.MTX"}}, 
                                            {{"Label", segments[2] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[2]+"_" +seg+ "_HH_Probabilities.MTX"}}, 
                                            {{"Label", segments[3] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[3]+"_" +seg+ "_HH_Probabilities.MTX"}}, 
                                            {{"Label", segments[4] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[4]+"_" +seg+ "_HH_Probabilities.MTX"}},  
                                            {{"Label", segments[5] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[5]+"_" +seg+ "_HH_Probabilities.MTX"}}, 
                                            {{"Label", segments[6] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[6]+"_" +seg+ "_HH_Probabilities.MTX"}}, 
                                            {{"Label", segments[7] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[7]+"_" +seg+ "_HH_Probabilities.MTX"}},  
                                            {{"Label", segments[8] +" - Probability"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Probabilities.mtx"}, {"File Name", output_path + segments[8]+"_" +seg+ "_HH_Probabilities.MTX"}}} 
    
    Opts.Output.[Applied Totals Matrices] = {{{"Label", segments[1] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName","Applied_Totals.MTX"}, {"File Name", output_path + segments[1]+"_" +seg+ "_HH_Applied_Totals.MTX"}}, 
                                            {{"Label", segments[2] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[2]+"_" +seg+ "_HH_Applied_Totals.MTX"}}, 
                                            {{"Label", segments[3] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[3]+"_" +seg+ "_HH_Applied_Totals.MTX"}}, 
                                            {{"Label", segments[4] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[4]+"_" +seg+ "_HH_Applied_Totals.MTX"}}, 
                                            {{"Label", segments[5] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[5]+"_" +seg+ "_HH_Applied_Totals.MTX"}}, 
                                            {{"Label", segments[6] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[6]+"_" +seg+ "_HH_Applied_Totals.MTX"}},  
                                            {{"Label", segments[7] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[7]+"_" +seg+ "_HH_Applied_Totals.MTX"}},  
                                            {{"Label", segments[8] +" - Applied Totals"}, {"Type", "Automatic"}, {"File based", "Automatic"}, {"Sparse", "Automatic"}, {"Column Major", "Automatic"}, {"Compression", 1}, {"FileName", "Applied_Totals.MTX"}, {"File Name", output_path + segments[8]+"_" +seg+ "_HH_Applied_Totals.MTX"}}}
    
    ok = RunMacro("TCB Run Procedure", "NestedLogitEngine", Opts, &Ret)
    
    endMacro

//============================================================================================
//  Finalize MC Outputs
//============================================================================================
    Macro "FinalizeHHMC" (path)
        purp = {"HBW", "WHBO", "HBOLOT", "HBOSOT"}
        suff = {"AS2", "AS1", "AS0"}
        ext = "_HH_Applied_Totals.MTX"
        per = {"PK", "OP","PK", "OP","PK", "OP","PK", "OP"}
        
        AS2cores ={"HBW_AGW_PK_AS2", "HBW_AGW_OP_AS2", "WHBO_AGW_PK_AS2", "WHBO_AGW_OP_AS2", "HBOLOT_AGW_PK_AS2", "HBOLOT_AGW_OP_AS2", "HBOSOT_AGW_PK_AS2", "HBOSOT_AGW_OP_AS2"}
        AS1cores ={"HBW_ALW_PK_AS1", "HBW_ALW_OP_AS1", "WHBO_ALW_PK_AS1", "WHBO_ALW_OP_AS1", "HBOLOT_ALW_PK_AS1", "HBOLOT_ALW_OP_AS1", "HBOSOT_ALW_PK_AS1", "HBOSOT_ALW_OP_AS1"}
        AS0cores ={"HBW_ZA_PK_AS0", "HBW_ZA_OP_AS0", "WHBO_ZA_PK_AS0", "WHBO_ZA_OP_AS0", "HBOLOT_ZA_PK_AS0", "HBOLOT_ZA_OP_AS0", "HBOSOT_ZA_PK_AS0", "HBOSOT_ZA_OP_AS0"}


        for i =1 to Arraylength(purp) do
           
                pur = purp[i]
                
                mcs = OpenMatrix(path + "CarShareAvailable.mtx", )
                mc = CreateMatrixCurrency(mcs, "CarShareAvail", "RCIndex", "RCIndex", )
                new_mat = CopyMatrix(mc, {{"File Name",path + "MC_OUT_" + pur +"_"+ per[i] + ".mtx"},{"Label", "MC Outputs"},{"Indices", "All"}})

                addcores = {"SOV", "HOV", "Transit", "Walk", "Bike"}
                
                m = OpenMatrix(path + "MC_OUT_" + pur +"_"+ per[i] + ".mtx", )

                for k=1 to ArrayLength(addcores) do
                    AddMatrixCore(m, addcores[k])
                end
                DropMatrixCore(m, "CarShareAvail")

                mc_array = CreateMatrixCurrencies(m, "TAZ", "TAZ", )

                m2 = OpenMatrix(path + AS2cores[i] + ext, )
                mc_array2 = CreateMatrixCurrencies(m2, "Origins", "Destinations", )

                m1 = OpenMatrix(path + AS1cores[i] +  ext, )
                mc_array1 = CreateMatrixCurrencies(m1, "Origins", "Destinations", )

                m0 = OpenMatrix(path + AS0cores[i] +  ext, )
                mc_array0 = CreateMatrixCurrencies(m0, "Origins", "Destinations", )


                // Summarize vehicle trips by mode
                mc_array.SOV := ((mc_array2.HHVEH-mc_array2.srHHVEH) + nz(mc_array2.ZIPcar)) + ((mc_array1.HHVEH-mc_array1.srHHVEH) + nz(mc_array1.ZIPcar)) + ((mc_array0.HHVEH-mc_array0.srHHVEH) + nz(mc_array0.ZIPcar))
                mc_array.HOV := (mc_array2.srHHVEH  + mc_array2.Carpool + mc_array2.Service) + (mc_array1.srHHVEH  + mc_array1.Carpool + mc_array1.Service) + (mc_array0.srHHVEH  + mc_array0.Carpool + mc_array0.Service)
                mc_array.Transit := mc_array2.Transit + mc_array1.Transit + mc_array0.Transit
                mc_array.Walk := mc_array2.Walk + mc_array1.Walk + mc_array0.Walk
                mc_array.Bike := mc_array2.Bike + mc_array1.Bike + mc_array0.Bike
            
        end

    endMacro

//============================================================================================
//  Finalize MC Outputs
//============================================================================================
    Macro "FinalizeUMC" (path)
        purpu = {"U_CBC", "U_HBC", "U_CBO", "U_HBO"}
        suffu = {"AS1", "AS0"}
        ext = "_HH_Applied_Totals.MTX"
        per = {"PK", "OP","PK", "OP","PK", "OP","PK", "OP"}

        UAS1cores ={"U_CBC_AS1_PK", "U_CBC_AS1_OP", "U_HBC_AS1_PK", "U_HBC_AS1_OP", "U_CBO_AS1_PK", "U_CBO_AS1_OP", "U_HBO_AS1_PK", "U_HBO_AS1_OP"}
        UAS0cores ={"U_CBC_AS0_PK", "U_CBC_AS0_OP", "U_HBC_AS0_PK", "U_HBC_AS0_OP", "U_CBO_AS0_PK", "U_CBO_AS0_OP", "U_HBO_AS0_PK", "U_HBO_AS0_OP"}

         for i =1 to Arraylength(purpu) do
                
                pur = purpu[i]

                mcs = OpenMatrix(path + "CarShareAvailable.mtx", )
                mc = CreateMatrixCurrency(mcs, "CarShareAvail", "RCIndex", "RCIndex", )
                new_mat = CopyMatrix(mc, {{"File Name",path + "MC_OUT_" + pur +"_"+ per[i] + ".mtx"},{"Label", "MC Outputs"},{"Indices", "All"}})

                addcores = {"SOV", "HOV", "Transit", "Walk", "Bike"}
                
                m = OpenMatrix(path + "MC_OUT_" + pur +"_"+ per[i] + ".mtx", )

                for k=1 to ArrayLength(addcores) do
                    AddMatrixCore(m, addcores[k])
                end
                DropMatrixCore(m, "CarShareAvail")

                mc_array = CreateMatrixCurrencies(m, "TAZ", "TAZ", )


                m1 = OpenMatrix(path + UAS1cores[i] + "_AS1" + ext, )
                mc_array1 = CreateMatrixCurrencies(m1, "Origins", "Destinations", )

                m0 = OpenMatrix(path + UAS0cores[i] + "_AS0" +ext, )
                mc_array0 = CreateMatrixCurrencies(m0, "Origins", "Destinations", )


                // Summarize vehicle trips by mode
                mc_array.SOV := ((mc_array1.HHVEH-mc_array1.srHHVEH) + nz(mc_array1.ZIPcar)) + ((mc_array0.HHVEH-mc_array0.srHHVEH) + nz(mc_array0.ZIPcar))
                mc_array.HOV := (mc_array1.srHHVEH  + mc_array1.Carpool + mc_array1.Service) + (mc_array0.srHHVEH  + mc_array0.Carpool + mc_array0.Service)
                mc_array.Transit := mc_array1.Transit + mc_array0.Transit
                mc_array.Walk := mc_array1.Walk + mc_array0.Walk
                mc_array.Bike := mc_array1.Bike + mc_array0.Bike

        end


    endMacro