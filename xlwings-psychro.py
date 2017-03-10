"""
Created on Fri Mar 10 2017

@author: gaevans

version: 0 revision 0


"""

import xlwings as xw

from math import atan
from math import exp
from math import log

# thermophysical constants
__Rda = 53.352                     # ft-lbf/(lb - R)
__Rv = 85.778                      # ft-lbf/(lb - R)
__cp_da = 0.240                    # Btu/(lbm - R)
__cp_w = 0.999                     # Btu/(lbm - R)
__cp_v = 0.451                     # Btu/(lbm - R)
__h_fg = 970.5                     # Btu/lbm
__molec_mass_ratio = 0.62198       # Mw/Mda (ratio of molecular masses for vapor & dry air)


# Twb - wet bulb temperature calculated from iterative method (deg F)
#
# Adapted from research listed in:
# Al-Ismaili, A.M., Al-Azri, N.A. (2016). Simple Iterative Approach to Calculate
# Wet-Bulb Temperature for Estimating Evaporative Cooling Efficiency
# International Journal of Agriculture Innovations and Research, 4(6), (Online) 1013-1018.
# http://ijair.org/administrator/components/com_jresearch/files/publications/IJAIR_1945_Final.pdf
# @param - dry bulb temperature (deg F)
# @param - relative humidity (%)
# @param - altitude (ft)
# @return - wet bulb temperature in degrees F

@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def Twb( Tdb, RH, Altitude ):
	
	Twb = Tdb                       # Initial guess at Twb
	W = humidity_ratio( Tdb, RH, Altitude )
	W_0 = W0( Tdb, Twb, RH, Altitude )
	epsilon = 1e-8                  # stop condition

	while ( abs( W_0 - W ) > epsilon ):
		if ( ( W_0 - W ) > 0 ):
			Twb = Twb * (1 - ( abs( W - W_0 ) / W ) / 100 )
			W_0 = W0( Tdb, Twb, RH, Altitude )
		elif ( ( W_0 - W ) < 0 ):
			Twb = Twb * (1 + ( abs( W - W_0 ) / W ) / 100 )
			W_0 = W0( Tdb, Twb, RH, Altitude )        
	return ( Twb ) 
# Twb_reg - wet bulb temperature regression calculation
#
# @param - dry bulb temperature (deg F)
# @param - relative humidity (%)
# @return - wet bulb temperature in degrees F

@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )

def Twb_reg( Tdb, RH ):
	
	Tdb = F_to_C( Tdb )
	# 'coefficients' of regression equation
	# Twb = Tdb * a( RH )  + b( Tdb, RH ) - c( RH ) + d( RH ) - e
	a = atan( ( 0.151977 * ( RH + 8.313659 ) ** 0.5 ) )
	bT = atan( Tdb + RH ) 
	c = atan( RH - 1.676331 )
	d = 0.00391838 * RH ** 1.5 * atan( 0.023101 * RH )
	e = 4.686035
	
	Twb = Tdb * a + bT - c + d - e
	Twb = C_to_F( Twb )
	
	return ( Twb )

# Tdp - dew point temperature (deg F)
#
# Regression calculation of Tdp from Peppers 1988 - ASHRAE 2001
# Equations 37 & 38
# @param - None
# @return - dew point temperature in degrees F
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )

def Tdp( Tdb, RH ):

	c = [ 100.45,
		  33.193,
		  2.319,
		  0.17074,
		  1.2063 ]
	
	P_w = Pw( Tdb, RH ) 			                # psia
	a = log( P_w )                      			# ln( psia )
	Tdp = 0                             			# deg C
	
	if ( Tdb >= 32.0 ):
		for i in xrange( len( c ) ):
			if ( i < 4 ):
				Tdp += c[ i ] * a ** i
			else:
				Tdp += c[ i ] * P_w ** ( 0.1984 )
	else:
		Tdp += 90.12 + 26.142 * a + 0.8927 * a ** 2.0

	return ( Tdp )

# relative_humidity - relative humidity (P/Pv)
#
# @param - dry bulb temperature in degrees F
# @return - relative humidity
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )

def relative_humidity( Tdb, RH ):
	P_w = Pw( Tdb, RH )
	P_ws = Pws( Tdb )
	
	return (P_w / P_ws)

# humidity_ratio_sat - saturated specific humidity or saturated Humidity Ratio (lbm/lbm [mv/ma])
#
# @param - dry bulb temperature in degrees F
# @return - saturation humidity ratio
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def humidity_ratio_sat( Tdb, RH, Altitude ):
	
	P_ws = Pws( Tdb )
	P = Patm_std( Altitude )
	MMR = __molec_mass_ratio

	return ( MMR * P_ws /( P - P_ws ) )

# humidity_ratio - specific humidity or Humidity Ratio (lbm/lbm)
#
# @param None
# @return humidity ratio (or specific humidity)
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def humidity_ratio( Tdb, RH, Altitude ):
	
	P = Patm_std( Altitude )
	P_ws = Pws( Tdb )
	MMR = __molec_mass_ratio
	RH = verify_RH( RH )
	#Pw = Pws * RH
	
	return ( MMR * ( P_ws * RH /( P - P_ws * RH) ) ) # equation is from ASHRAE 2002 / 2005
	#return ( 0.62198 * Pw /( P - Pw * 0.378) ) # need citation for this equation - from wikipedia, but no citation
	
# humidity ratio - calculation as a function of Tdb and Twb with variable Cp for air, water, and vapor
#
# Adaped from research listed in:
# Al-Ismaili, A.M., Al-Azri, N.A. (2016). Simple Iterative Approach to Calculate
# Wet-Bulb Temperature for Estimating Evaporative Cooling Efficiency
# International Journal of Agriculture Innovations and Research, 4(6), (Online) 1013-1018.
# http://ijair.org/administrator/components/com_jresearch/files/publications/IJAIR_1945_Final.pdf
#
# @param - dry bulb temperature in degrees F
# @param - wet bulb temperature in degrees F
# @return - humidity ratio

def W0 ( Tdb, Twb, RH, Altitude ):
	
	h_fg = __h_fg              # Latent Heat of Vaporization (Btu/lbm)
	Cp_w = Cp_water            # Cp of water function
	Cp_da = Cp_dry_air         # Cp of dry air function
	Cp_v = Cp_vapor            # Cp of water vapor function
	Ws_wb = humidity_ratio_sat # saturated humidity ratio Ws(Twb)
	z = Altitude
	
	a = ( ( h_fg - ( Cp_w( Twb ) - Cp_v( Tdb ) ) * Twb ) * Ws_wb( Twb, RH, z ) - Cp_da( Tdb, Twb ) * ( Tdb - Twb ) )
	b = ( h_fg + Cp_v( Tdb ) * Tdb - Cp_w( Twb ) * Twb )
	
	return ( a / b )

# humidity ratio - calculation as a function of Tdb and Twb with const specific heats for air, water, vapor
#
# Adaped from research listed in:
# Al-Ismaili, A.M., Al-Azri, N.A. (2016). Simple Iterative Approach to Calculate
# Wet-Bulb Temperature for Estimating Evaporative Cooling Efficiency
# International Journal of Agriculture Innovations and Research, 4(6), (Online) 1013-1018.
# http://ijair.org/administrator/components/com_jresearch/files/publications/IJAIR_1945_Final.pdf
#
# @param - dry bulb temperature in degrees F
# @param - wet bulb temperature in degrees F
# @return - humidity ratio based on constant specific heat of air, water, and vapor

def W1 ( Tdb, Twb, RH, Altitude ):
	
	h_fg = __h_fg              # Latent Heat of Vaporization (Btu/lbm)        
	Cp_da = __cp_da            # specific heat of dry air
	Cp_w = __cp_w              # specific heat of water 
	Cp_v = __cp_v              # specific heat of vapor 
	Ws_wb = humidity_ratio_sat # saturated humidity ratio Ws(Twb)
	z = Altitude
	
	a = ( ( h_fg - ( Cp_w - Cp_v ) * Twb ) * Ws_wb( Twb, RH, z ) - Cp_da * ( Tdb - Twb ) )
	b = ( h_fg + Cp_v * Tdb - Cp_w * Twb )
	
	return ( a / b )

# abs_humidity - absolute numidity (lbm/ft^3 [mv/V])
#
# @param None
# @return - absolute humidity
@xw.func
def abs_humidity(  ):
	pass

# deg_sat( ) - returns the degree of saturation of the binary mixture (dry air / vapor)
# 
# @param None
# @return - degree of saturation (quality)
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def deg_sat( Tdb, RH, Altitude ):
	RH = verify_RH( RH )
	W = humidity_ratio( Tdb, RH, Altitude )
	Ws = humidity_ratio_sat( Tdb, RH, Altitude)
	
	return ( W / Ws )

# specific_vol - specific volume (ft^3 / lbm dry air)
#
# @param None
# @return - specific volume of air/vapor mixture
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def specific_vol( Tdb, RH, Altitude ):
	P = Patm_std( Altitude )                   				# psia
	Tdb_R = F_to_R( Tdb )          							# deg R
	RH = verify_RH( RH )
	W = humidity_ratio( Tdb, RH, Altitude )			# lbm/lbm
	
	return ( 0.3704 * Tdb_R * ( 1 + 1.6078 * W ) / P )
	
		
# enthalpy - specific enthalpy of moist air (Btu / lbm dry air)
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def enthalpy( Tdb, RH, Altitude ):
	RH = verify_RH( RH )
	W = humidity_ratio( Tdb, RH , Altitude)
	
	return ( __cp_da * Tdb + W * (1060.9 + 0.435 * Tdb) )

# entropy - specific entropy of moist air (Btu / lbm dry air)
@xw.func
def entropy ( ):
	pass

# Pw - partial pressure for vapor at the specified dry bulb temperature
# @param None
# @return partial pressure Pv (psia)
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'RH' , doc = 'relative humidity' )
    
def Pw( Tdb, RH ):
	P_ws = Pws( Tdb )
	RH = verify_RH( RH )
	
	return ( RH * P_ws )
	
# Pws - returns the water vapor saturation pressure for temperatures
# SATURATION VAPOR PRESSURE IN ABSENCE OF DRY AIR
# in the range of -148 to 392 deg F
# @param T - temperature (deg F)
# @param units - unit of measure for temperature, default is deg F
# @return - saturation pressure (psi absolute)
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )

def Pws( Tdb ):
	
	T_abs_K = C_to_K( F_to_C( Tdb ) )
	
	# calculate vapor pressure over ice
	if ( T_abs_K <= 273.15 and T_abs_K >= 173.15 ):
		return ( P_sat_ice_HW( T_abs_K ) )
	# calculate vapor pressure over water
	elif ( T_abs_K > 273.15 and T_abs_K <= 473.15 ):
		return ( P_sat_water_HW( T_abs_K ) )
	
# __P_sat_ice - returns the water vapor saturation pressure over ice (psia)
# Hyland-Wexler Correlations - 1983 - ASHRAE 2001
#
# @param T_abs - Thermodynamics temperature ( K )
# @return saturation vapor pressure over ice (psia)

def P_sat_ice_HW( T_abs ):
   m = [ -0.56745359e04,
		 -.51523058,
		 -0.009677843,
		 0.62215701e-6,
		 0.20747825e-08,
		 -0.94840240e-12,
		 0.41635019e01]
   
   # ln(Pw) = Sum(m_i * T^i-1, i = 0, 5) + m_6 * ln(T)
   ln_P_sat_ice = 0
   for i in xrange( -1, len( m ) - 1 ):
	   if ( i < len( m ) - 2 ):
		   ln_P_sat_ice += m[ i + 1 ] * T_abs ** i
	   else:
		   ln_P_sat_ice += m[ i + 1] * log( T_abs )
		   
   Pws_ice = exp( ln_P_sat_ice )
			
   return ( Pa_to_psia( Pws_ice ) )
	
# __P_sat_water - returns the water vapor saturation pressure over water (psia)
# Hyland-Wexler Correlations - 1983 - ASHRAE 2001
#
# @param T_abs - Thermodynamic temperature ( K )
# @return saturation vapor pressure over water (psia)	

def P_sat_water_HW( T_abs ):
	
	h = [ -0.58002206e04,
		  -5.516256,
		  -0.48640239e-01,
		  0.41764768e-04,
		  -0.14452093e-07,
		  6.5459673e00 ]

   # ln(Pw,ice) = Sum(h_i * T^1, i = -1, 3) + h_r * ln(T)        
	ln_P_sat_water = 0
	for i in xrange( -1, len( h ) - 1 ):
		if ( i < len( h ) - 2 ):
			ln_P_sat_water += h[ i + 1 ] * T_abs ** i
		else:
			ln_P_sat_water += h[ i + 1 ] * log( T_abs )
	   
	Pws_water = exp( ln_P_sat_water )

	return ( Pa_to_psia( Pws_water ) )

# Cp_water - specific heat capacity of water at constant pressure (Btu/lbm-R)
# Modified to convert from SI to imperial units
#
# Adapted from research listed in:
# A. M. Al-Ismaili, Modelling of a humidification
# dehumidification greenhouse in Oman. Cranfield University, 2009, pp 190-191.
#
# K. Raznjevic, Handbook of thermodynamic tables and charts
# Washington: Hemisphere Publishing Corp., 1975 pp. 67-68.
#
# @param - wet bulb temperature in degrees F
# @return - specific heat of water in Btu/lbm-F
@xw.func
@xw.arg( 'Twb', doc = 'wet bulb temperature (deg F)' )

def Cp_water( Twb ):
	
	Twb = F_to_C( Twb )
	# convert from J/kg-C to kJ/kg-C
	Cp_w = ( 0.0265 * Twb ** 2 - 1.7688 * Twb + 4205.6 ) / 1000.0
	
	return ( kJ_kgK_2_Btu_lbF( Cp_w ) )
# Cp_vapor - specific heat capacity of water vapor in dry air (Btu/lbm-R)
# Modified to convert from SI to imperial units
#
# Adapted from research listed in:
# A. M. Al-Ismaili, Modelling of a humidification
# dehumidification greenhouse in Oman. Cranfield University, 2009, pp 190-191.
#
# K. Raznjevic, Handbook of thermodynamic tables and charts
# Washington: Hemisphere Publishing Corp., 1975 pp. 67-68.
#
# @param - dry bulb temperature in degrees F
# @return - specific heat of water vapor in Btu/lbm-F
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )

def Cp_vapor( Tdb ):
	
	Tdb = F_to_C( Tdb )
	# convert from J/kg-C to kJ/kg-C
	Cp_vapor = ( 0.0016 * Tdb ** 2 + 0.1546 * Tdb + 1858.7 ) / 1000.0
	
	return ( kJ_kgK_2_Btu_lbF( Cp_vapor ) )

# Cp_dry_air - specific heat capacity of dry air (Btu/lbm-R)
# Modified to convert from SI to imperial units
#
# Adapted from research listed in:
# A. M. Al-Ismaili, Modelling of a humidification
# dehumidification greenhouse in Oman. Cranfield University, 2009, pp 190-191.
#
# K. Raznjevic, Handbook of thermodynamic tables and charts
# Washington: Hemisphere Publishing Corp., 1975 pp. 67-68.
#
# @param - dry bulb temperature in degrees F 
# @param - wet bulb temperature in degrees F
# @return - specific head of dry air in Btu/lbm-F
@xw.func
@xw.arg( 'Tdb', doc = 'dry bulb temperature (deg F)' )
@xw.arg( 'Twb', doc = 'wet bulb temperature (deg F)' )

def Cp_dry_air( Tdb, Twb ):
	
	Twb = F_to_C( Twb )
	Tdb = F_to_C( Tdb )
	
	# convert from J/kg-C to kJ/kg-C
	Cp_dry_air = ( 0.0667 * ( Tdb + Twb ) / 2 + 1005 ) / 1000.0
	
	return ( kJ_kgK_2_Btu_lbF( Cp_dry_air ) )	

# Patm_std - standard pressure at a given altitude (psia)
# @param None
# @return - standard atmospheric pressure based on altitude (ft)
@xw.func
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def Patm_std( Altitude ):
	if ( is_valid_altitude( Altitude ) ):
		return ( 14.696 * ( 1 - Altitude * 6.8754e-06 ) ** 5.2559 )
	
# Tatm_std - standard temperature at a given altitude in degrees F
# @param None
# @return - standard temperature based on altitude (deg F)
@xw.func
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def Tatm_std( Altitude ):
	if ( is_valid_altitude( Altitude ) ):
		T = 59 - 0.00356620 * Altitude
	return ( T )
	
@xw.func
@xw.arg( 'Altitude', doc = 'local altitude (ft)' )

def is_valid_altitude( Altitude ):
	if ( Altitude >= -1000 and Altitude <= 30000 ):
		return True
	return False

def verify_RH( RH ):
	if ( RH >= 1 and RH <= 100 ):
		return RH / 100
	elif ( RH > 0 and RH < 1 ):
		return RH

# F_to_R - returns the conversion of temperature from farenheit to rankine
@xw.func
def F_to_R( T_farenheit ):
    return ( T_farenheit + 459.67 )

# C_to_K - returns the Thermodynamic temperature from deg C to K
@xw.func
def C_to_K( T_celsius ):
    return ( T_celsius + 273.15 )

# F_to_C - returns the conversion of temperature from farenheit to celsius
@xw.func
def F_to_C( T_farenheit ):
    return ( ( T_farenheit - 32.0 ) * 5.0 / 9.0 )

# C_to_F - returns the conversion of temperature from celsius to farenheit
@xw.func
def C_to_F( T_celsius ):
    return ( T_celsius * ( 9.0 / 5.0 ) + 32.0 )

# K_to_C - returns the conversion of temperature from Kelvin to celsius
@xw.func
def K_to_C( T_kelvin ):
    return ( T_kelvin - 273.15 )

# R_to_F - returns the conversion of temperature from Rankine to farenheit
@xw.func
def R_to_F( T_rankine ):
    return ( T_rankine - 459.67 )
	
@xw.func	
def Pa_to_psia( P ):
    return ( P * 1.45038e-01 )
@xw.func
def psia_to_Pa( P ):
    return ( P * 6894.76 )
@xw.func
def kPa_to_psia( P ):
    return ( Pa_to_psia( P / 101.325 ) )
@xw.func
def bar_to_psia( P ):
    return ( P * 1e3 )
@xw.func
def mBar_to_psia( P ):
    return ( P * 0.145038 )

#------------------------------------------------------------------------------------------------------------------------------
# MISC Conversions
@xw.func
def Btu_lbF_2_kJ_kgK( val ):
    return ( val * 4.1868 )           # kJ/kg-K
	
@xw.func
def kJ_kgK_2_Btu_lbF( val ):
    return ( val * 0.238846 )         # Btu/lb-F

@xw.func
def feet_2_meters( val ):
	return ( val * 0.3048 )

@xw.func
def meters_2_feet( val ):
	return ( val * 3.28084 )



