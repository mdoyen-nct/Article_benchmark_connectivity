function outputFile = compute_classification_metrics(inputFile, outputFile)
%COMPUTE_CLASSIFICATION_METRICS Export binary classifier metrics to Excel.
%   COMPUTE_CLASSIFICATION_METRICS() reads the "all" worksheet from
%   CN_vs_AD_byCenter.xlsx and writes classification_metrics.xlsx.
%
%   The positive class is Group == 1. Rows with a missing grouping value,
%   ground truth, or classifier result are excluded from the corresponding
%   calculation.

scriptDirectory = fileparts(mfilename('fullpath'));
if nargin < 1 || isempty(inputFile)
    inputFile = fullfile(scriptDirectory, 'CN_vs_AD_byCenter.xlsx');
end
if nargin < 2 || isempty(outputFile)
    outputFile = fullfile(scriptDirectory, 'classification_metrics.xlsx');
end

data = readtable(inputFile, 'Sheet', 'all');

ageName = findVariable(data, 'subject_age');
siteName = findVariable(data, 'SITEID');
petName = findVariable(data, 'PET');
groupName = findVariable(data, 'Group');
classifierColumns = {'Classif DPI', 'Classif PCA', 'Classif SICE'};
classifierNames = ["DPI", "PCA", "SICE"];

truth = asNumeric(data.(groupName), groupName);
age = asNumeric(data.(ageName), ageName);
predictions = NaN(height(data), numel(classifierColumns));
for classifierIndex = 1:numel(classifierColumns)
    columnName = findVariable(data, classifierColumns{classifierIndex});
    predictions(:, classifierIndex) = asNumeric(data.(columnName), columnName);
end

validateBinary(truth, groupName);
for classifierIndex = 1:numel(classifierNames)
    validateBinary(predictions(:, classifierIndex), classifierNames(classifierIndex));
end

site = strtrim(string(data.(siteName)));
pet = strtrim(string(data.(petName)));
site(ismissing(data.(siteName))) = missing;
pet(ismissing(data.(petName))) = missing;
ageDecadeStart = floor(age / 10) * 10;
ageDecade = strings(height(data), 1);
validAge = ~isnan(ageDecadeStart);
ageDecade(validAge) = compose('%d-%d', ageDecadeStart(validAge), ...
    ageDecadeStart(validAge) + 9);

bySite = summarizeGroups(site, 'SiteID', truth, predictions, classifierNames);
byPET = summarizeGroups(pet, 'PET', truth, predictions, classifierNames);
byAge = summarizeGroups(ageDecade, 'AgeDecade', truth, predictions, classifierNames);
overall = summarizeGroups(repmat("All", height(data), 1), 'Population', ...
    truth, predictions, classifierNames);
statisticalTests = fitMixedEffectsModels(site, pet, ageDecade, truth, ...
    predictions, classifierNames);
statisticalNotes = table( ...
    {'Model'; 'Outcome'; 'FixedEffects'; 'RandomEffect'; 'Tests'; ...
    'AdjustedPValue'; 'Significant'; 'Inclusion'; ...
    'Interpretation'}, ...
    {'Mixed-effects logistic regression fitted separately for each classifier'; ...
    'Classification correctness (correct=1, incorrect=0)'; ...
    'PET and age decade'; ...
    'Random intercept for center (SiteID)'; ...
    'Global marginal F-test for each fixed effect'; ...
    'Benjamini-Hochberg correction across all tests'; ...
    'AdjustedPValue < 0.05'; ...
    'All group sizes are included; only rows missing a model variable are excluded'; ...
    'An association does not prove a causal effect; site and PET machine may be confounded'}, ...
    'VariableNames', {'Item', 'Description'});

if isfile(outputFile)
    delete(outputFile);
end
writeResults(bySite, outputFile, 'BySite');
writeResults(byPET, outputFile, 'ByPET');
writeResults(byAge, outputFile, 'ByAgeDecade');
writeResults(overall, outputFile, 'Overall');
writeResults(statisticalTests, outputFile, 'StatisticalTests');
writetable(statisticalNotes, outputFile, 'Sheet', 'StatisticalNotes');

fprintf('Results written to %s\n', outputFile);
end

function variableName = findVariable(data, requestedName)
names = string(data.Properties.VariableNames);
normalizedNames = lower(regexprep(names, '[^a-zA-Z0-9]', ''));
normalizedRequestedName = lower(regexprep(string(requestedName), ...
    '[^a-zA-Z0-9]', ''));
match = find(normalizedNames == normalizedRequestedName, 1);
if isempty(match)
    error('Missing required column "%s".', requestedName);
end
variableName = data.Properties.VariableNames{match};
end

function values = asNumeric(values, variableName)
if isnumeric(values) || islogical(values)
    values = double(values);
else
    values = str2double(strtrim(string(values)));
end
values = values(:);
if all(isnan(values))
    error('Column "%s" does not contain numeric values.', variableName);
end
end

function validateBinary(values, variableName)
observed = unique(values(~isnan(values)));
if any(observed ~= 0 & observed ~= 1)
    error('Column "%s" must contain only 0, 1, or missing values.', variableName);
end
end

function result = summarizeGroups(groupValues, groupColumnName, truth, ...
        predictions, classifierNames)
groupValues = groupValues(:);
validGroup = ~ismissing(groupValues) & strlength(groupValues) > 0;
groups = unique(groupValues(validGroup), 'sorted');
numberOfRows = numel(groups) * numel(classifierNames);

groupOutput = strings(numberOfRows, 1);
classifierOutput = strings(numberOfRows, 1);
n = zeros(numberOfRows, 1);
tp = zeros(numberOfRows, 1);
tn = zeros(numberOfRows, 1);
fp = zeros(numberOfRows, 1);
fn = zeros(numberOfRows, 1);
sensitivity = NaN(numberOfRows, 1);
specificity = NaN(numberOfRows, 1);
accuracy = NaN(numberOfRows, 1);
ppv = NaN(numberOfRows, 1);
balancedAccuracy = NaN(numberOfRows, 1);
npv = NaN(numberOfRows, 1);
positiveLikelihoodRatio = NaN(numberOfRows, 1);
negativeLikelihoodRatio = NaN(numberOfRows, 1);

row = 0;
for groupIndex = 1:numel(groups)
    inGroup = validGroup & groupValues == groups(groupIndex);
    for classifierIndex = 1:numel(classifierNames)
        row = row + 1;
        prediction = predictions(:, classifierIndex);
        valid = inGroup & ~isnan(truth) & ~isnan(prediction);
        actual = truth(valid);
        estimated = prediction(valid);

        groupOutput(row) = groups(groupIndex);
        classifierOutput(row) = classifierNames(classifierIndex);
        n(row) = sum(valid);
        tp(row) = sum(actual == 1 & estimated == 1);
        tn(row) = sum(actual == 0 & estimated == 0);
        fp(row) = sum(actual == 0 & estimated == 1);
        fn(row) = sum(actual == 1 & estimated == 0);

        sensitivity(row) = safeDivide(tp(row), tp(row) + fn(row));
        specificity(row) = safeDivide(tn(row), tn(row) + fp(row));
        accuracy(row) = safeDivide(tp(row) + tn(row), n(row));
        ppv(row) = safeDivide(tp(row), tp(row) + fp(row));
        npv(row) = safeDivide(tn(row), tn(row) + fn(row));
        balancedAccuracy(row) = (sensitivity(row) + specificity(row)) / 2;
        positiveLikelihoodRatio(row) = sensitivity(row) / (1 - specificity(row));
        negativeLikelihoodRatio(row) = (1 - sensitivity(row)) / specificity(row);
    end
end

result = table(groupOutput, classifierOutput, n, tp, tn, fp, fn, ...
    sensitivity, specificity, accuracy, ppv, balancedAccuracy, npv, ...
    positiveLikelihoodRatio, negativeLikelihoodRatio, ...
    'VariableNames', {groupColumnName, 'Classifier', 'N', 'TP', 'TN', ...
    'FP', 'FN', 'Sensitivity', 'Specificity', 'Accuracy', 'PPV', ...
    'BalancedAccuracy', 'NPV', 'PositiveLikelihoodRatio', ...
    'NegativeLikelihoodRatio'});
end

function value = safeDivide(numerator, denominator)
if denominator == 0
    value = NaN;
else
    value = numerator / denominator;
end
end

function writeResults(result, outputFile, sheetName)
% Excel displays NaN values as empty cells; use NA to mark undefined metrics.
for columnIndex = 1:width(result)
    values = result{:, columnIndex};
    if isnumeric(values) && any(isnan(values))
        outputValues = num2cell(values);
        outputValues(isnan(values)) = {'NA'};
        result.(result.Properties.VariableNames{columnIndex}) = outputValues;
    end
end
writetable(result, outputFile, 'Sheet', sheetName);
end

function result = fitMixedEffectsModels(site, pet, ageDecade, truth, ...
        predictions, classifierNames)
factorNames = ["PET", "AgeDecade"];
numberOfRows = numel(factorNames) * numel(classifierNames);

factorOutput = strings(numberOfRows, 1);
classifierOutput = strings(numberOfRows, 1);
n = zeros(numberOfRows, 1);
numberOfCenters = zeros(numberOfRows, 1);
numberOfLevels = zeros(numberOfRows, 1);
fStatistic = NaN(numberOfRows, 1);
numeratorDegreesOfFreedom = NaN(numberOfRows, 1);
denominatorDegreesOfFreedom = NaN(numberOfRows, 1);
pValue = NaN(numberOfRows, 1);

row = 0;
validModelVariables = ~ismissing(site) & strlength(site) > 0 & ...
    ~ismissing(pet) & strlength(pet) > 0 & ...
    ~ismissing(ageDecade) & strlength(ageDecade) > 0 & ~isnan(truth);
for classifierIndex = 1:numel(classifierNames)
    prediction = predictions(:, classifierIndex);
    valid = validModelVariables & ~isnan(prediction);
    modelData = table(categorical(site(valid)), categorical(pet(valid)), ...
        categorical(ageDecade(valid)), prediction(valid) == truth(valid), ...
        'VariableNames', {'Center', 'PET', 'AgeDecade', 'Correct'});
    model = fitglme(modelData, ...
        'Correct ~ PET + AgeDecade + (1|Center)', ...
        'Distribution', 'Binomial', 'Link', 'Logit');
    modelTests = anova(model);

    for factorIndex = 1:numel(factorNames)
        row = row + 1;
        testRow = find(string(modelTests.Properties.RowNames) == ...
            factorNames(factorIndex), 1);
        if isempty(testRow)
            error('Mixed-effects model did not return a test for %s.', ...
                factorNames(factorIndex));
        end

        factorOutput(row) = factorNames(factorIndex);
        classifierOutput(row) = classifierNames(classifierIndex);
        n(row) = height(modelData);
        numberOfCenters(row) = numel(categories(modelData.Center));
        if factorNames(factorIndex) == "PET"
            numberOfLevels(row) = numel(categories(modelData.PET));
        else
            numberOfLevels(row) = numel(categories(modelData.AgeDecade));
        end
        fStatistic(row) = modelTests.FStat(testRow);
        numeratorDegreesOfFreedom(row) = modelTests.DF1(testRow);
        denominatorDegreesOfFreedom(row) = modelTests.DF2(testRow);
        pValue(row) = modelTests.pValue(testRow);
    end
end

adjustedPValue = benjaminiHochberg(pValue);
significant = adjustedPValue < 0.05;
result = table(factorOutput, classifierOutput, n, numberOfCenters, ...
    numberOfLevels, fStatistic, numeratorDegreesOfFreedom, ...
    denominatorDegreesOfFreedom, pValue, adjustedPValue, significant, ...
    'VariableNames', {'Factor', 'Classifier', 'N', 'NumberOfCenters', ...
    'NumberOfLevels', 'FStatistic', 'NumeratorDegreesOfFreedom', ...
    'DenominatorDegreesOfFreedom', 'PValue', 'AdjustedPValue', ...
    'Significant'});
end

function adjusted = benjaminiHochberg(probabilities)
adjusted = NaN(size(probabilities));
valid = ~isnan(probabilities);
[sortedProbabilities, order] = sort(probabilities(valid));
numberOfTests = numel(sortedProbabilities);
if numberOfTests == 0
    return;
end

sortedAdjusted = sortedProbabilities .* numberOfTests ./ (1:numberOfTests)';
for index = numberOfTests - 1:-1:1
    sortedAdjusted(index) = min(sortedAdjusted(index), ...
        sortedAdjusted(index + 1));
end
sortedAdjusted = min(sortedAdjusted, 1);
validIndices = find(valid);
adjusted(validIndices(order)) = sortedAdjusted;
end
