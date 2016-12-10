/*
**************************************
* Validação de tipos v1.0            *
* Autor: Wagner B. Soares            *
**************************************
*/

isNull = function(x)
{
	if((x == 'undefined') || (x == null)){return true;}
	else{return false;}
};

isObject = function(x)
{
	if(!isNull(x))
	{
		if(x.constructor == Object){return true;}
		else{return false;}
	}
	else{return false;}
};

isFunction = function(x)
{
	if(!isNull(x))
	{
		if(x instanceof Function){return true;}
		else{return false;}
	}
	else{return false;}
}

isBoolean = function(x)
{
	if(!isNull(x))
	{
		if(x.constructor == Boolean){return true;}
		else{return false;}
	}
	else{return false;}
};

isArray = function(x)
{
	if(!isNull(x))
	{
		if(x.constructor == Array){return true;}
		else{return false;}
	}
	else{return false;}
};

isString = function(x)
{
	if(!isNull(x))
	{
		if(x.constructor == String){return true;}
		else{return false;}
	}
	else{return false;}
};

isDate = function(x)
{
	if(!isNull(x))
	{
		if(x.constructor == Date){return true;}
		else{return false;}
	}
	else{return false;}
};

isNumber = function(x)
{
	if(!isNull(x))
	{
		if(!isNaN(x) && (x.constructor != Boolean) && (x.constructor != Array)){return true;}
		else{return false;}
	}
	else{return false;}
};

isInteger = function(x)
{
	if(!isNull(x))
	{
		if(isNumber(x))
		{
			if((x%1) == 0){return true;}
			else{return false;}
		}
		else{return false;}
	}
	else{return false;}
};