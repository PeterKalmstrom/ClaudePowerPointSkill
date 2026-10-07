"""Audit every Python function in the skill for the error pattern in scripts/kShared.py.

    python tools/check_kpattern.py            # all of scripts/, tools/, mcp-app/
    python tools/check_kpattern.py FILE ...

Per function (methods, module functions, MCP tools):
  1 guard     first statement (after the docstring) is `if kS.ErrorMode:` ending in return/raise
              (not required in __init__ or in kS/kRun, the infrastructure)
  2 try       the rest of the body is one try statement
  3 report    its last handler catches Exception and calls kS.GlobalErrorHandler(e, "<Class>.<Method>...")
              with a literal location naming this function; earlier handlers may only re-raise
  4 default   that handler ends in return or raise
Also banned anywhere: lambda, nested functions (use a named method). A function may opt out with a comment
"DOCUMENTED EXCEPTION: <why>" on its def line or the line above (pure helpers whose callers are all guarded).
Exit 1 when anything is missing; each finding names file, line and function.
"""
import ast
import glob
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INFRA = {"kS", "kRun", "kErrorDetail", "kErrorReport", "kToolException", "ToolInputException", "ToolReportableException"}


class kPatternAudit:
    """Walks one file's syntax tree and collects every function that breaks the pattern."""

    def __init__(self, Path):
        self.Path = Path
        self.Rel = os.path.relpath(Path, ROOT)
        self.Lines = open(Path, encoding="utf-8").read().splitlines()
        self.Tree = ast.parse("\n".join(self.Lines), Path)
        self.Module = os.path.splitext(os.path.basename(Path))[0]
        self.Findings = []

    def Add(self, Node, Message):
        self.Findings.append(f"{self.Rel}:{Node.lineno}: {getattr(Node, 'name', 'lambda')}: {Message}")

    def Exempt(self, Node):
        for Line in (Node.lineno - 1, Node.lineno - 2):
            if 0 <= Line < len(self.Lines) and "DOCUMENTED EXCEPTION:" in self.Lines[Line]:
                return True
        return False

    def Run(self):
        for Node in ast.walk(self.Tree):
            if isinstance(Node, ast.Lambda):
                self.Findings.append(f"{self.Rel}:{Node.lineno}: lambda: use a named method")
        for Node in self.Tree.body:
            self.Visit(Node, None)
        return self.Findings

    def Visit(self, Node, ClassName):
        if isinstance(Node, ast.ClassDef):
            for Child in Node.body:
                self.Visit(Child, Node.name)
        elif isinstance(Node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            self.CheckFunction(Node, ClassName)

    def CheckFunction(self, Node, ClassName):
        for Inner in ast.walk(Node):
            if Inner is not Node and isinstance(Inner, (ast.FunctionDef, ast.AsyncFunctionDef)):
                self.Add(Inner, "nested function: make it a named method")
        if ClassName in INFRA or self.Exempt(Node):
            return
        Body = list(Node.body)
        if Body and isinstance(Body[0], ast.Expr) and isinstance(getattr(Body[0], "value", None), ast.Constant) \
                and isinstance(Body[0].value.value, str):
            Body = Body[1:]
        if Node.name != "__init__":
            if not Body or not self.IsGuard(Body[0]):
                self.Add(Node, "1 guard missing: first statement must be `if kS.ErrorMode: return <default>`")
                return
            Body = Body[1:]
        if len(Body) != 1 or not isinstance(Body[0], ast.Try):
            self.Add(Node, "2 the whole body (after the guard) must be one try statement")
            return
        self.CheckHandlers(Node, ClassName, Body[0])

    @staticmethod
    def IsGuard(Stmt):
        return (isinstance(Stmt, ast.If) and ast.unparse(Stmt.test) == "kS.ErrorMode" and not Stmt.orelse
                and isinstance(Stmt.body[-1], (ast.Return, ast.Raise)))

    def CheckHandlers(self, Node, ClassName, Try):
        if not Try.handlers:
            self.Add(Node, "3 try has no except Exception handler")
            return
        Last = Try.handlers[-1]
        if Last.type is None or ast.unparse(Last.type) not in ("Exception", "BaseException"):
            self.Add(Node, "3 the last handler must catch Exception")
            return
        for Earlier in Try.handlers[:-1]:
            if not (len(Earlier.body) == 1 and isinstance(Earlier.body[0], ast.Raise) and Earlier.body[0].exc is None):
                self.Add(Node, f"3 handler for {ast.unparse(Earlier.type) if Earlier.type else 'all'} may only re-raise")
        Calls = [c for c in ast.walk(ast.Module(body=Last.body, type_ignores=[]))
                 if isinstance(c, ast.Call) and ast.unparse(c.func) == "kS.GlobalErrorHandler"]
        if not Calls:
            self.Add(Node, "3 the Exception handler must call kS.GlobalErrorHandler")
            return
        Args = Calls[0].args
        Want = f"{ClassName or self.Module}.{Node.name}"
        if len(Args) < 2 or not isinstance(Args[1], (ast.Constant, ast.JoinedStr)):
            self.Add(Node, f'3 location must be a literal "{Want}"')
        else:
            Text = Args[1].value if isinstance(Args[1], ast.Constant) else ast.unparse(Args[1])[2:]
            if not str(Text).startswith(Want):
                self.Add(Node, f'3 location "{Text}" does not name this function ("{Want}")')
        if Node.name != "__init__" and not isinstance(Last.body[-1], (ast.Return, ast.Raise)):
            self.Add(Node, "4 the Exception handler must end in return <safe default> or raise")


class kPatternAuditApp:
    """Command line: audit the given files, or every Python file in the skill."""

    def Run(self):
        Files = sys.argv[1:] or sorted(glob.glob(os.path.join(ROOT, "scripts", "*.py"))
                                       + glob.glob(os.path.join(ROOT, "tools", "*.py"))
                                       + glob.glob(os.path.join(ROOT, "mcp-app", "*.py")))
        Findings = []
        for Path in Files:
            Findings += kPatternAudit(Path).Run()
        for Line in Findings:
            print(Line)
        print(f"{len(Files)} file(s) checked, {len(Findings)} problem(s)")
        return 1 if Findings else 0


if __name__ == "__main__":  # the auditor stays standard-library only and outside the pattern it checks
    sys.exit(kPatternAuditApp().Run())
