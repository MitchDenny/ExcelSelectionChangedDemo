function PreviousSelection(status, jsonText, sequence) {
    var self = this;
    this.status = ko.observable(status);
    this.jsonText = ko.observable(jsonText);
    this.sequence = ko.observable(sequence);
}

function MainViewModel() {
    var self = this;
    this.previousSelections = ko.observableArray();
}

Office.initialize = function (reason) {
    $(document).ready(function () {

        var history = document.getElementById('History');
        viewModel = new MainViewModel();
        ko.applyBindings(viewModel, history);

        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            SelectionChanged
            );
    });
};

function SelectionChanged(args) {
    args.document.getSelectedDataAsync(Office.CoercionType.Matrix, {}, DataSelected);
}

function DataSelected(result) {
    var sequence = viewModel.previousSelections().length + 1;
    var status = result.status;
    var jsonText = JSON.stringify(result, {}, ' ');
    var previousSelection = new PreviousSelection(status, jsonText, sequence);
    viewModel.previousSelections.unshift(previousSelection);
}