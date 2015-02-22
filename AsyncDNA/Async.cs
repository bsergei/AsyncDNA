using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using AsyncDNA.Integration;
using Excel;
using ExcelDna.Integration;
using Reference = AsyncDNA.Integration.Reference;

namespace AsyncDNA
{
    /// <summary>
    /// Service that exposes the only Calc function for async calculations.
    /// Uses ExcelDNA's RTD asynchronous implementation under the hood.
    /// </summary>
    public class Async : IDisposable
    {
        private readonly IXlService xlService_;
        private readonly ILogService logService_;
        
        /// <summary>
        /// Value that means function execution is postponed while dependency is calculating.
        /// </summary>
        public const string Scheduled = "Scheduled...";

        /// <summary>
        /// Value that means function is doing its background work and result isn't available yet.
        /// </summary>
        public const string Calculating = "Calculating...";
        
        private readonly HashSet<string> knownAsyncFuncs_ = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

        // Evaluation state variables.
        private readonly HashSet<ReferenceFunc> calculatedFunctions_ = new HashSet<ReferenceFunc>();
        private readonly Dictionary<ExcelReference, AsyncCallInfo> calculatingRefs_ = new Dictionary<ExcelReference, AsyncCallInfo>();
        private readonly Dictionary<ExcelReference, EvalState> evalStates_ = new Dictionary<ExcelReference, EvalState>();

        private DateTime lastFinished_ = DateTime.MinValue;
        private DateTime batchStarted_ = DateTime.MinValue;
        private long batchId_;

        public Async(IXlService xlService, ILogService logService)
        {
            xlService_ = xlService ?? new NullXlService();
            logService_ = logService ?? new NullLogService();
            SetThrottleInterval(0); // Boost RTD processing and update interval.
        }

        public void Dispose()
        {
            SetThrottleInterval(2000); // Return to default value.
        }

        private static void SetThrottleInterval(int value)
        {
            dynamic app = ExcelDnaUtil.Application;
            app.RTD.ThrottleInterval = value;
        }

        /// <summary>
        /// Adds function name to well-known async funcs lists.
        /// </summary>
        public void RegisterAsyncFunc(string funcName)
        {
            knownAsyncFuncs_.Add(funcName);
        }
        
        /// <summary>
        /// Starts async UDF calculation.
        /// Returns Scheduled/Calculating or calculated value.
        /// Make sure calling UDF marked with IsMacroType=true.
        /// </summary>
        public object Calc(CallArguments args)
        {
            CheckFuncName(args);
            MarkUdfNonVolatile();
            ExcelReference caller;
            try
            {
                caller = GetCallerExcelReference();
            }
            catch (Exception e)
            {
                logService_.WriteError(
                    "Error getting reference. " +
                    FunctionLoggerHelper.LogParameters(args.FunctionName, null), e);
                // TODO Cleanup
                return "Error getting reference.";
            }

            if (IsNewEvaluationBatchStarted)
                OnEvaluatingBatchStarted();

            object[] asyncKeyArgs = GetAsyncKeyArgs(args.RawArgs, caller, batchStarted_);
            AsyncCallInfo asyncCallInfo = new AsyncCallInfo(args.FunctionName, asyncKeyArgs);

            if (!IsCalculatingTheSame(caller, asyncCallInfo) || IsScheduled(caller))
            {
                // Check simple case. Precedent is still calculating.
                Precedents precedents = caller.GetPrecedents();
                if (precedents.GetRefs().Any(IsCalculatingAny))
                {
                    SetEvalState(caller, EvalState.Scheduled, args, asyncCallInfo);
                    return Scheduled;
                }

                if (ForcePrecedentCalculated(caller, precedents))
                {
                    SetEvalState(caller, EvalState.Scheduled, args, asyncCallInfo);
                    return Scheduled;
                }
            }

            Validate(args);

            bool isCalculating;
            object result = Observe(asyncKeyArgs, caller.XlfToReference(), args, out isCalculating);

            if (isCalculating)
            {
                // Async call to service scheduled.
                SetEvalState(caller, EvalState.ObservableCreated, args, asyncCallInfo);
                return Calculating;
            }
            else
            {
                // Async call to service and result are ready.
                SetEvalState(caller, EvalState.ObservableFinished, args, asyncCallInfo);
                return result;
            }
        }

        private bool ForcePrecedentCalculated(ExcelReference caller, Precedents precedents)
        {
            // Ensure that all precedents have been calculated, force calculation otherwise.
            foreach (ReferenceFunc rf in precedents.GetFormulas())
            {
                string funcName = rf.FuncName;
                if (knownAsyncFuncs_.Contains(funcName)) // Check that precedent is our known function, so we force its calculation if need.
                {
                    // We just check that our function has been calculated at the specified reference.
                    if (!calculatedFunctions_.Contains(rf))
                    {
                        ExcelReference precedent = rf.ExcelReference;

                        bool shouldForceCalculation = IsPrecedentCalcShouldForced(caller, precedent, funcName);
                        if (shouldForceCalculation)
                        {
                            if (!evalStates_.ContainsKey(precedent))
                            {
                                // Force evaluate precedent because we don't know its state.
                                ExcelAsyncUtil.QueueAsMacro(() => ForceFormulaEvaluate(precedent));
                            }
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Can be overriden in subclasses to make a decision, whether precedent should be forced to calculate again.
        /// This is necessary, e.g., in case, when application function returns a handle to pre-calculated structure,
        /// which has not been yet calculated or should be calculated each time when dependant is evaluating.
        /// Returns false by default.
        /// </summary>
        protected virtual bool IsPrecedentCalcShouldForced(ExcelReference caller, ExcelReference precedent, string precedentFuncName)
        {
            return false;
        }

        private static void Validate(CallArguments args)
        {
            if (args.RawArgs.Any(_ => Calculating.Equals(_)))
                Debug.Fail("'Calculating' parameter in function call");

            if (args.RawArgs.Any(_ => Scheduled.Equals(_)))
                Debug.Fail("'Scheduled' parameter in function call");
        }

        private static void MarkUdfNonVolatile()
        {
            // Async functions should be non-volatile only.
            #region #8.6.5 in Financial Applications Using Excel Add-in Development in C/C++

            /* Functions registered as macro-sheet equivalents, type #, and as taking xloper or
            xloper12 arguments, type R and U, rather than the value-only types P and Q, are by
            default volatile. This echoes the behaviour of XLM macro sheets when the ARGUMENT()
            function was used with the parameter 8 to specify that a given argument should be left
            as a reference. The logic behind Excel treating these functions as volatile is that if you
            want to calculate something based on the reference, i.e. the location of a cell, then you
            must recalculate every time in case the location has changed but the value has stayed the
            same.
            It is possible to alter the volatile status of an XLL function with a call to the C
            API function xlfVolatile, passing a Boolean false xloper/xloper12 argument.
            However, there are reports that this can confuse Excel’s order-of-recalculation logic, so
            the advice would be to decide at the outset whether your functions need to be volatile or
            not, and stick with that. */

            #endregion
            XlCall.Excel(XlCall.xlfVolatile, false);
        }

        private void CheckFuncName(CallArguments args)
        {
            if (!knownAsyncFuncs_.Contains(args.FunctionName))
                throw new Exception("Function is not registered to be async");
        }

        private static ExcelReference GetCallerExcelReference()
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return caller;
        }

        private static object[] GetAsyncKeyArgs(object[] rawArgs, ExcelReference excelReference, DateTime batchTimeStamp)
        {
            var rawArgsWithRef = new object[rawArgs.Length + 2];
            Array.Copy(rawArgs, rawArgsWithRef, rawArgs.Length);
            
            // Add Excel reference to input args. This will prevent caching result value for each two cells having the same input args.
            rawArgsWithRef[rawArgs.Length] = excelReference;

            // Need to add batch time stamp to prevent caching: ExcelDNA cache resulting value by input parameters.
            rawArgsWithRef[rawArgs.Length + 1] = batchTimeStamp.ToOADate(); 
            return rawArgsWithRef;
        }

        /// <summary>
        /// Ensure this function is not running in UDF context. It should be run in macro context.
        /// </summary>
        private void ForceFormulaEvaluate(ExcelReference excelReference)
        {
            string formula = (string) XlCall.Excel(XlCall.xlfGetCell, 6, excelReference);
            Range comRange = excelReference.ToComRange();
            if (comRange.HasArray)
            {
                comRange.CurrentArray.FormulaArray = formula;
            }
            else
            {
                comRange.Formula = formula;
            }
        }

        private void RemoveFromCalculating(ExcelReference caller)
        {
            calculatingRefs_.Remove(caller);
            for (int row = caller.RowFirst; row <= caller.RowLast; row++)
            {
                for (int col = caller.ColumnFirst; col <= caller.ColumnLast; col++)
                {
                    ExcelReference excelReference = new ExcelReference(row, row, col, col, caller.SheetId);
                    calculatingRefs_.Remove(excelReference);
                }
            }
        }

        private void MarkAsCalculating(ExcelReference caller, AsyncCallInfo asyncCallInfo)
        {
            calculatingRefs_[caller] = asyncCallInfo;
            for (int row = caller.RowFirst; row <= caller.RowLast; row++)
            {
                for (int col = caller.ColumnFirst; col <= caller.ColumnLast; col++)
                {
                    ExcelReference excelReference = new ExcelReference(row, row, col, col, caller.SheetId);
                    calculatingRefs_[excelReference] = asyncCallInfo;
                }
            }
        }

        private bool IsCalculatingTheSame(ExcelReference excelReference, AsyncCallInfo asyncCallInfo)
        {
            AsyncCallInfo cached;
            return calculatingRefs_.TryGetValue(excelReference, out cached) && cached.Equals(asyncCallInfo);
        }

        private bool IsCalculatingAny(ExcelReference excelReference)
        {
            AsyncCallInfo stored;
            return calculatingRefs_.TryGetValue(excelReference, out stored);
        }

        private bool IsScheduled(ExcelReference excelReference)
        {
            EvalState evalState;
            return evalStates_.TryGetValue(excelReference, out evalState) && evalState == EvalState.Scheduled;
        }

        private void CleanupLastEvaluatingBatch()
        {
            calculatedFunctions_.Clear();
            calculatingRefs_.Clear();
            evalStates_.Clear();
        }

        private bool IsEvaluatingBatchSeemsToBeFinished
        {
            get { return evalStates_.Values.All(_ => _ == EvalState.ObservableFinished); }
        }

        private void OnEvaluatingBatchStarted()
        {
            CleanupLastEvaluatingBatch();
            batchStarted_ = DateTime.UtcNow;
            batchId_++;
        }

        private void LogState()
        {
            string msg;
            if (IsEvaluatingBatchSeemsToBeFinished)
            {
                msg = "";
            }
            else
            {
                int calculating = evalStates_.Values.Count(_ => _ == EvalState.ObservableCreated);
                int calculated = evalStates_.Values.Count(_ => _ == EvalState.ObservableFinished);
                int scheduled =
                    evalStates_.Values.Count(_ => _ == EvalState.Scheduled || _ == EvalState.ForcingFormulaUpdate);

                msg = String.Format("{0}: Evaluating async funcs ({1} calculating, {2} calculated, {3} scheduled)...",
                    batchId_,
                    calculating,
                    calculated,
                    scheduled);
            }
            ExcelAsyncUtil.QueueAsMacro(() => XlCall.Excel(XlCall.xlcMessage, true, msg));
        }

        private void UpdateLastFinished()
        {
            lastFinished_ = DateTime.UtcNow;
        }

        private bool IsNewEvaluationBatchStarted
        {
            get
            {
                if (!IsEvaluatingBatchSeemsToBeFinished)
                    return false;

                dynamic app = ExcelDnaUtil.Application;
                int throttleInterval = app.RTD.ThrottleInterval;
                int batchInterval = throttleInterval + 500;
                double lastFinishedMilliseconds = (DateTime.UtcNow - lastFinished_).TotalMilliseconds;
                return lastFinishedMilliseconds > batchInterval;
            }
        }

        private void SetEvalState(ExcelReference caller, EvalState evalState, CallArguments args, AsyncCallInfo asyncCallInfo)
        {
            evalStates_[caller] = evalState;
            switch (evalState)
            {
                case EvalState.ObservableCreated:
                case EvalState.Scheduled:
                case EvalState.ForcingFormulaUpdate:
                    MarkAsCalculating(caller, asyncCallInfo);
                    LogState();
                    break;

                case EvalState.ObservableFinished:
                    RemoveFromCalculating(caller);
                    AddCalculatedFunc(caller, args);

                    if (IsEvaluatingBatchSeemsToBeFinished)
                        UpdateLastFinished();

                    LogState();
                    break;
            }
        }

        private void AddCalculatedFunc(ExcelReference caller, CallArguments args)
        {
            string funcName = args.FunctionName;
            AddCalculatedFormulaArray(caller, funcName);
        }

        private void AddCalculatedFormulaArray(ExcelReference caller, string funcName)
        {
            for (int row = caller.RowFirst; row <= caller.RowLast; row++)
            {
                for (int col = caller.ColumnFirst; col <= caller.ColumnLast; col++)
                {
                    ExcelReference arrayRef = new ExcelReference(row, row, col, col, caller.SheetId);
                    calculatedFunctions_.Add(new ReferenceFunc(arrayRef, funcName));
                }
            }
        }

        private object Observe(object[] rawArgsWithRef, Reference reference, CallArguments args, out bool isCalculating)
        {
            IXlService xlService = xlService_;
            ILogService logService = logService_;

            object result = RxExcel.Observe(args.FunctionName, rawArgsWithRef,
                () => new CallObservable(reference, args, xlService, logService));

            isCalculating = result is ExcelError && ((ExcelError) result) == ExcelError.ExcelErrorNA;
            if (isCalculating)
            {
                return null;
            }
            else
            {
                return result;
            }
        }
    }
}