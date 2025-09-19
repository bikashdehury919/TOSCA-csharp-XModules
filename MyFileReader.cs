using System.IO;

using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using Tricentis.Automation.Interaction.SpecialExecutionTasks;

namespace AutomationExtensions;

[SpecialExecutionTaskName("MyFileReader")]
internal class MyFileReader : SpecialExecutionTaskEnhanced {

    public MyFileReader(Validator validator)
        : base(validator) { }

    public override void ExecuteTask(ISpecialExecutionTaskTestAction testAction) {

        IParameter path;
        IParameter content;

       // MainConfiguration.Instance.TryGet(“SynchronizationTimeout”, out string  value);

        // First read the parameters passed to SETs.
        // TBox provides SetParameterValidator which allows you to do basic validations on your SET parameters. 
        // For example for Path parameter we are dictating that the it is required and cannot have an empty value, and also the action mode can be only Input:
        try {
            var validator = new SetParameterValidator(testAction);
            path = validator.Take("Path").Required().NotEmpty().Accepts(new[]{ActionMode.Input}).Run();
            content = validator.Take("Content").Required().Accepts(new[]{ActionMode.Verify, ActionMode.Buffer, ActionMode.Input}).Run();
        } catch (ParameterValidationException) {
            // in case any of the criteria set for parameters didn't meet a ParameterValidationException will be raised,
            // and in that case you should not continue with the rest of execution.
            return;
        }

        var filePath = path.ValueAsString();
        if (!File.Exists(filePath)) {
            testAction.SetResult(new UnknownFailedActionResult("File not found"));
            return;
        }

        if (content.ActionMode == ActionMode.Input) {
            var contentValue = content.ValueAsString();
            File.WriteAllText(filePath, contentValue);
        } else {
            var fileValue = File.ReadAllText(filePath);
            // HandleActualValue handles the Verify, Buffer, and WaitOn action modes
            HandleActualValue(testAction, content, fileValue);
        }
    }
}
