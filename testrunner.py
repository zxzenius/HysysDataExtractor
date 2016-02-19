# -*- coding: utf-8 -*-
"""Only for test service
"""

import win32com.client

__author__ = 'zenius'


def extract_data(filename):
    hy_app = win32com.client.gencache.EnsureDispatch('HYSYS.Application.V7.3')
    hy_app.Visible = False
    hy_case = hy_app.SimulationCases.Open(filename)
    main_streams = hy_case.Flowsheet.MaterialStreams


def get_streamprop(streams):
    for stream in streams:
        print(stream.name)


def get_node(stream):
    nodeprop = dict()
    nodeprop['ActualGasFlow'] = stream.ActualGasFlow()
    nodeprop['ActualLiqFlow'] = stream.ActualLiqFlow()
    nodeprop['ActualVolumeFlow'] = stream.ActualVolumeFlow()
    nodeprop['AvgLiqDensity'] = stream.AvgLiqDensity()
    nodeprop['BOBubblePointPressure'] = stream.BOBubblePointPressure()
    nodeprop['BOBubblePointTemperature'] = stream.BOBubblePointTemperature()
    nodeprop['Stream.BOGasOilRatio'] = stream.BOGasOilRatio()
    nodeprop['BOMassEnthalpy'] = stream.BOMassEnthalpy()
    nodeprop['BOMassFlow'] = stream.BOMassFlow()
    nodeprop[
        'BOOilFormationVolumeFactor'] = stream.BOOilFormationVolumeFactor()
    nodeprop['BOOilViscosity'] = stream.BOOilViscosity()
    nodeprop['BOPressure'] = stream.BOPressure()
    nodeprop['BOSolutionGOR'] = stream.BOSolutionGOR()
    nodeprop['BOSpecificGravity'] = stream.BOSpecificGravity()
    nodeprop['BOSurfaceTension'] = stream.BOSurfaceTension()
    # Failed in under line
    # nodeprop['BOTemperatureInVM'] = stream.BOTemperatureInVM()
    nodeprop['BOTemperature'] = stream.BOTemperature()
    nodeprop['BOViscosityCoefficientA'] = stream.BOViscosityCoefficientA()
    nodeprop['BOViscosityCoefficientB'] = stream.BOViscosityCoefficientB()
    # Failed in this line
    # nodeprop['BOViscosity'] = stream.BOViscosity()
    nodeprop['BOVolumetricFlow'] = stream.BOVolumetricFlow()
    nodeprop['BOWaterCut'] = stream.BOWaterCut()
    nodeprop['BOWatsonK'] = stream.BOWatsonK()
    nodeprop['ComponentMassFlow'] = stream.ComponentMassFlow()
    nodeprop['ComponentMolarFlow'] = stream.ComponentMolarFlow()
    nodeprop['ComponentMolarFraction'] = stream.ComponentMolarFraction()
    nodeprop['ComponentVolumeFlow'] = stream.ComponentVolumeFlow()
    nodeprop['ComponentVolumeFraction'] = stream.ComponentVolumeFraction()
    nodeprop['Compressibility'] = stream.Compressibility()
    nodeprop['CpCv'] = stream.CpCv()
    nodeprop['EnthalpyEstimate'] = stream.EnthalpyEstimate()
    nodeprop['FlowEstimate'] = stream.FlowEstimate()
    nodeprop['HeatFlow'] = stream.HeatFlow()
    nodeprop['HeatOfVap'] = stream.HeatOfVap()
    nodeprop['HeavyLiquidFraction'] = stream.HeavyLiquidFraction()
    nodeprop['HigherHeatValue'] = stream.HigherHeatValue()
    nodeprop['IdealLiquidVolumeFlow'] = stream.IdealLiquidVolumeFlow()
    nodeprop['IsEnergyStream'] = stream.IsEnergyStream
    nodeprop['IsValid'] = stream.IsValid
    nodeprop['KineticViscosity'] = stream.KineticViscosity()
    nodeprop['LightLiquidFraction'] = stream.LightLiquidFraction()
    nodeprop['LiquidFraction'] = stream.LiquidFraction()
    nodeprop['LowerHeatValue'] = stream.LowerHeatValue()
    nodeprop['MassDensity'] = stream.MassDensity()
    nodeprop['MassEnthalpy'] = stream.MassEnthalpy()
    nodeprop['MassEntropy'] = stream.MassEntropy()
    nodeprop['MassFlow'] = stream.MassFlow()
    nodeprop['MassHeatCapacity'] = stream.MassHeatCapacity()
    nodeprop['MassHeatOfVap'] = stream.MassHeatOfVap()
    nodeprop['MassHigherHeatValue'] = stream.MassHigherHeatValue()
    nodeprop['MassLowerHeatValue'] = stream.MassLowerHeatValue()
    nodeprop['MolarDensity'] = stream.MolarDensity()
    nodeprop['MolarEnthalpy'] = stream.MolarEnthalpy()
    nodeprop['MolarEntropy'] = stream.MolarEntropy()
    nodeprop['MolarFlow'] = stream.MolarFlow()
    nodeprop['MolarHeatCapacity'] = stream.MolarHeatCapacity()
    nodeprop['MolarVolume'] = stream.MolarVolume()
    nodeprop['MolecularWeight'] = stream.MolecularWeight()
    nodeprop['name'] = stream.name
    nodeprop['Power'] = stream.Power()
    nodeprop['PressureCO2'] = stream.PressureCO2()
    nodeprop['Pressure'] = stream.Pressure()
    nodeprop['SGAir'] = stream.SGAir()
    nodeprop['StdGasFlow'] = stream.StdGasFlow()
    nodeprop['StdLiqMassDensity'] = stream.StdLiqMassDensity()
    nodeprop['StdLiqVolFlow'] = stream.StdLiqVolFlow.Value
    nodeprop['StreamDescription'] = stream.StreamDescription
    nodeprop['SurfaceTension'] = stream.SurfaceTension()
    nodeprop['TaggedName'] = stream.TaggedName
    nodeprop['TemperatureEstimate'] = stream.TemperatureEstimate()
    nodeprop['TemperatureValue'] = stream.Temperature()
    nodeprop['ThermalConductivity'] = stream.ThermalConductivity()
    nodeprop['TypeName'] = stream.TypeName
    nodeprop['UniqueID'] = stream.UniqueID
    nodeprop['VapourFraction'] = stream.VapourFraction()
    nodeprop['VisibleTypeName'] = stream.VisibleTypeName
    nodeprop['WatsonK'] = stream.WatsonK()
    # Catch FluidPackageName for ComponentList
    nodeprop['FluidPackage'] = stream.FluidPackage()
    return nodeprop

def get_fluidpkgs(hy_case):
    fluidpkgs = dict()
    for FluidPackage in hy_case.BasisManager.FluidPackages:
        fluidpkg = dict()
        fluidpkg['PropertyPackageName'] = FluidPackage.PropertyPackageName
        fluidpkg['Components'] = FluidPackage.ComponentList.Components.Names
        fluidpkgs[FluidPackage()] = fluidpkg
    return fluidpkgs

def get_streams(hy_case):
    streams_db = dict()
    # Add SubFlows to List
    flowsheets = list(hy_case.Flowsheet.Flowsheets)
    flowsheets.append(hy_case.Flowsheet)
    for flowsheet in flowsheets:
        for stream in flowsheet.MaterialStreams:
            streams_db[stream.TaggedName] = get_node(stream)
    return streams_db
