"""
Equation Parser Module
Handles parsing and evaluation of user-defined pricing equations.
Supports variables like TB (table price), WD (with damper price), and mathematical operations.
Supports model references like [MODEL] which resolve to model prices at the same dimensions.
"""

import re
import math
from typing import Dict, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from utils.price_calculator import PriceCalculator


class EquationParser:
    """Parses and evaluates pricing equations with variable substitution"""
    
    # Regex pattern for matching [MODEL] tokens in equations
    _MODEL_TOKEN_RE = re.compile(r'\[([A-Za-z0-9_-]+)\]')
    
    @staticmethod
    def is_number(value) -> bool:
        """Check if a value is a simple number (not an equation)"""
        if value is None:
            return False
        try:
            float(str(value).strip())
            return True
        except ValueError:
            return False
    
    @staticmethod
    def is_equation(value) -> bool:
        """Check if a value is an equation (contains variables or functions)"""
        if value is None:
            return False
        value_str = str(value).strip()
        # Check if it contains equation-like patterns
        # Note: we also treat bracketed tokens like "[WSD]" as equation-like so callers
        # can route them through the equation pipeline (they are expanded before eval).
        equation_indicators = ['TB', 'WD', 'SIZE', '[', ']', '(', ')', '+', '-', '*', '/', 'sqrt', 'max', 'min', 'round', 'abs', 'ceil', 'floor', 'pow']
        return any(indicator in value_str for indicator in equation_indicators)
    
    def parse_equation(self, equation: str, variables: Dict[str, float], 
                       price_calculator: Optional['PriceCalculator'] = None) -> float:
        """
        Parse and evaluate a pricing equation with variable substitution.
        
        Args:
            equation: The equation string (e.g., "(TB + WD)*4*1.45" or "TB+[WD]")
            variables: Dictionary of variable values (e.g., {'TB': 100, 'WD': 200, 'WIDTH': 8, 'HEIGHT': 60})
            price_calculator: Optional PriceCalculator instance to resolve [MODEL] tokens.
                            If provided, [MODEL] tokens in the equation will be expanded to numeric values.
        
        Returns:
            The calculated result
            
        Raises:
            ValueError: If equation is invalid or contains unsafe operations
        """
        if not equation or not equation.strip():
            raise ValueError("Empty equation")
        
        # Clean the equation
        equation = equation.strip()
        
        # Expand [MODEL] tokens if a price calculator is provided
        if price_calculator is not None:
            equation = self.expand_model_tokens(equation, variables, price_calculator)
        
        # Validate equation contains only allowed characters
        if not self._is_safe_equation(equation):
            raise ValueError("Equation contains invalid characters or operations")
        
        # Replace variables with their values
        substituted_equation = self._substitute_variables(equation, variables)
        
        # Create a safe evaluation environment
        safe_globals = {
            "__builtins__": {},
            # Add all math functions directly
            "sqrt": math.sqrt,
            "pow": math.pow,
            "ceil": math.ceil,
            "floor": math.floor,
            "abs": abs,
            "round": round,
            "min": min,
            "max": max,
            "sin": math.sin,
            "cos": math.cos,
            "tan": math.tan,
            "log": math.log,
            "log10": math.log10,
            "exp": math.exp,
            "pi": math.pi,
            "e": math.e,
        }
        
        # Evaluate the mathematical expression
        try:
            result = eval(substituted_equation, safe_globals, {})
            return float(result)
        except Exception as e:
            raise ValueError(f"Error evaluating equation: {str(e)}")
    
    def _is_safe_equation(self, equation: str) -> bool:
        """Check if equation contains only safe characters and operations"""
        if not equation or not equation.strip():
            return False
        
        # Clean the equation
        equation = equation.strip()
        
        # Check for obviously dangerous patterns
        dangerous_patterns = [
            r'import\s+',           # import statements
            r'from\s+',              # from imports
            r'__import__',           # __import__ calls
            r'exec\s*\(',            # exec() calls
            r'eval\s*\(',            # eval() calls
            r'open\s*\(',            # file operations
            r'file\s*\(',            # file() calls
            r'input\s*\(',           # input() calls
            r'raw_input\s*\(',       # raw_input() calls
            r'compile\s*\(',         # compile() calls
            r'globals\s*\(',         # globals() calls
            r'locals\s*\(',          # locals() calls
            r'vars\s*\(',            # vars() calls
            r'dir\s*\(',             # dir() calls
            r'getattr\s*\(',         # getattr() calls
            r'setattr\s*\(',         # setattr() calls
            r'hasattr\s*\(',         # hasattr() calls
            r'delattr\s*\(',         # delattr() calls
            r'__.*__',               # double underscore methods
            r'\.__.*__',             # double underscore attributes
            r'[;\n\r]',              # semicolons and newlines
            r'#.*',                  # comments
            r'""".*"""',             # triple quotes
            r"'''.*'''",             # triple quotes
        ]
        
        for pattern in dangerous_patterns:
            if re.search(pattern, equation, re.IGNORECASE | re.DOTALL):
                return False
        
        # Check for balanced parentheses
        if not self._has_balanced_parentheses(equation):
            return False
        
        # Check for basic mathematical syntax
        if not self._is_valid_math_syntax(equation):
            return False
        
        return True
    
    def _has_balanced_parentheses(self, equation: str) -> bool:
        """Check if parentheses are balanced"""
        count = 0
        for char in equation:
            if char == '(':
                count += 1
            elif char == ')':
                count -= 1
                if count < 0:
                    return False
        return count == 0
    
    def _is_valid_math_syntax(self, equation: str) -> bool:
        """Check if equation has valid mathematical syntax"""
        try:
            # Try to compile the expression to check syntax
            compile(equation, '<string>', 'eval')
            return True
        except SyntaxError:
            return False
        except Exception:
            return False
    
    def expand_model_tokens(self, equation: str, variables: Dict[str, float],
                           price_calculator: 'PriceCalculator') -> str:
        """
        Replace occurrences of [MODEL] in an equation with the referenced MODEL's TB price
        at the same dimensions as the current product.
        
        Args:
            equation: The equation string potentially containing [MODEL] tokens
            variables: Dictionary of variable values, should contain WIDTH/HEIGHT or DIAMETER for dimension context
            price_calculator: PriceCalculator instance to resolve model tokens to prices
        
        Returns:
            Equation string with [MODEL] tokens replaced by numeric values
            
        Dimension context is read from variables:
        - WIDTH, HEIGHT (default table)
        - DIAMETER (other table)
        """
        if not equation:
            return equation
        
        width = variables.get('WIDTH')
        height = variables.get('HEIGHT')
        diameter = variables.get('DIAMETER')
        
        def repl(m):
            model = m.group(1)
            tb = price_calculator._resolve_model_tb_price(model, width, height, diameter)
            return str(tb)
        
        # Use .sub() to replace all [MODEL] tokens in one pass
        return self._MODEL_TOKEN_RE.sub(repl, equation)
    
    def _substitute_variables(self, equation: str, variables: Dict[str, float]) -> str:
        """Replace variable names with their values in the equation"""
        substituted = equation
        
        # Replace variables with their values
        for var_name, var_value in variables.items():
            # Use word boundaries to ensure we don't replace partial variable names
            pattern = r'\b' + re.escape(var_name) + r'\b'
            substituted = re.sub(pattern, str(var_value), substituted)
        
        return substituted
