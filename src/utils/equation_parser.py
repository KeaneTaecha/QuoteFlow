"""
Equation Parser Module
Handles parsing and evaluation of user-defined pricing equations.
Supports variables like TB (table price), WD (with damper price), and mathematical operations.
"""

import re
import math
from typing import Dict


class EquationParser:
    """Parses and evaluates pricing equations with variable substitution"""
    
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
        equation_indicators = ['TB', 'WD', 'SIZE', '(', ')', '+', '-', '*', '/', 'sqrt', 'max', 'min', 'round', 'abs', 'ceil', 'floor', 'pow']
        return any(indicator in value_str for indicator in equation_indicators)
    
    def parse_equation(self, equation: str, variables: Dict[str, float]) -> float:
        """
        Parse and evaluate a pricing equation with variable substitution.
        
        Args:
            equation: The equation string (e.g., "(TB + WD)*4*1.45")
            variables: Dictionary of variable values (e.g., {'TB': 100, 'WD': 200})
        
        Returns:
            The calculated result
            
        Raises:
            ValueError: If equation is invalid or contains unsafe operations
        """
        if not equation or not equation.strip():
            raise ValueError("Empty equation")
        
        # Clean the equation
        equation = equation.strip()
        
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
    
    def _substitute_variables(self, equation: str, variables: Dict[str, float]) -> str:
        """Replace variable names with their values in the equation"""
        substituted = equation
        
        # Replace variables with their values
        for var_name, var_value in variables.items():
            # Use word boundaries to ensure we don't replace partial variable names
            pattern = r'\b' + re.escape(var_name) + r'\b'
            substituted = re.sub(pattern, str(var_value), substituted)
        
        return substituted
