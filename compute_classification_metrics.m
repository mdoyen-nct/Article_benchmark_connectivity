function outputFile = compute_classification_metrics(inputFile, outputFile, numberOfBootstrapSamples)
%COMPUTE_CLASSIFICATION_METRICS Export binary classifier metrics to Excel.
%   COMPUTE_CLASSIFICATION_METRICS() reads the "all" worksheet from
%   CN_vs_AD_byCenter.xlsx and writes classification_metrics.xlsx.
%   A parametric-bootstrap likelihood-ratio test (1,000 samples by default)
%   tests site heterogeneity after adjustment for age.
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
if nargin < 3 || isempty(numberOfBootstrapSamples)
    numberOfBootstrapSamples = 1000;
end
validateattributes(numberOfBootstrapSamples, {'numeric'}, ...
    {'scalar', 'integer', 'positive'}, mfilename, 'numberOfBootstrapSamples');

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
mixedEffectsModels = fitSiteEffectModels(site, age, truth, predictions, ...
    classifierNames, numberOfBootstrapSamples);
statisticalNotes = table( ...
    {'Models'; 'Outcome'; 'FixedEffects'; 'RandomEffects'; 'NullModel'; ...
    'Bootstrap'; 'MultipleTesting'; 'Inclusion'; 'ReportedResults'; ...
    'Interpretation'}, ...
    {'One mixed-effects logistic regression fitted for each classifier'; ...
    'Classification correctness (correct=1, incorrect=0)'; ...
    'Centered continuous age'; ...
    'Site random intercept; PET is not included in the inferential model'; ...
    'Logistic regression with centered continuous age and no site random effect'; ...
    sprintf('%d outcomes simulated under the null model; P values are unavailable if any fit fails', numberOfBootstrapSamples); ...
    'Raw bootstrap P values are adjusted across classifiers with the Holm method'; ...
    'Rows missing age, site, outcome, or classifier result are excluded'; ...
    'Site variance, ICC, likelihood-ratio statistic, bootstrap P value, and Holm-adjusted P value'; ...
    'A significant result indicates site-associated heterogeneity after age adjustment, not a causal effect'}, ...
    'VariableNames', {'Item', 'Description'});

if isfile(outputFile)
    delete(outputFile);
end
writeResults(bySite, outputFile, 'BySite');
writeResults(byPET, outputFile, 'ByPET');
writeResults(byAge, outputFile, 'ByAgeDecade');
writeResults(overall, outputFile, 'Overall');
writeResults(mixedEffectsModels, outputFile, 'MixedEffectsModels');
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

function result = fitSiteEffectModels(site, age, truth, predictions, ...
        classifierNames, numberOfBootstrapSamples)
numberOfRows = numel(classifierNames);

classifierOutput = classifierNames(:);
n = zeros(numberOfRows, 1);
numberOfLevels = zeros(numberOfRows, 1);
intercept = NaN(numberOfRows, 1);
ageCoefficient = NaN(numberOfRows, 1);
randomEffectVariance = NaN(numberOfRows, 1);
randomEffectStandardDeviation = NaN(numberOfRows, 1);
intraclassCorrelation = NaN(numberOfRows, 1);
nullLogLikelihood = NaN(numberOfRows, 1);
siteLogLikelihood = NaN(numberOfRows, 1);
likelihoodRatio = NaN(numberOfRows, 1);
bootstrapSamplesSuccessful = zeros(numberOfRows, 1);
bootstrapPValue = NaN(numberOfRows, 1);
bootstrapMonteCarloSE = NaN(numberOfRows, 1);

previousRandomState = rng;
restoreRandomState = onCleanup(@() rng(previousRandomState));
rng(0, 'twister');

for classifierIndex = 1:numel(classifierNames)
    prediction = predictions(:, classifierIndex);
    valid = ~ismissing(site) & strlength(site) > 0 & ~isnan(age) & ...
        ~isnan(truth) & ~isnan(prediction);
    centeredAge = age(valid) - mean(age(valid));
    modelData = table(categorical(site(valid)), centeredAge, ...
        prediction(valid) == truth(valid), ...
        'VariableNames', {'Site', 'AgeCentered', 'Correct'});

    nullModel = fitglm(modelData, 'Correct ~ 1 + AgeCentered', ...
        'Distribution', 'binomial', 'Link', 'logit');
    siteModel = fitglme(modelData, ...
        'Correct ~ 1 + AgeCentered + (1|Site)', ...
        'Distribution', 'Binomial', 'Link', 'Logit', ...
        'FitMethod', 'Laplace');
    covariance = covarianceParameters(siteModel);
    variance = covariance{1}(1, 1);

    n(classifierIndex) = height(modelData);
    numberOfLevels(classifierIndex) = numel(categories(modelData.Site));
    intercept(classifierIndex) = siteModel.Coefficients.Estimate(1);
    ageCoefficient(classifierIndex) = siteModel.Coefficients.Estimate(2);
    randomEffectVariance(classifierIndex) = variance;
    randomEffectStandardDeviation(classifierIndex) = sqrt(variance);
    intraclassCorrelation(classifierIndex) = variance / (variance + pi^2 / 3);
    nullLogLikelihood(classifierIndex) = nullModel.LogLikelihood;
    siteLogLikelihood(classifierIndex) = ...
        siteModel.ModelCriterion.LogLikelihood;
    likelihoodRatio(classifierIndex) = max(0, 2 * ...
        (siteLogLikelihood(classifierIndex) - ...
        nullLogLikelihood(classifierIndex)));

    nullProbability = predict(nullModel, modelData);
    bootstrapStatistics = NaN(numberOfBootstrapSamples, 1);
    for bootstrapIndex = 1:numberOfBootstrapSamples
        bootstrapData = modelData;
        bootstrapData.Correct = rand(height(modelData), 1) < nullProbability;
        try
            bootstrapNullModel = fitglm(bootstrapData, ...
                'Correct ~ 1 + AgeCentered', ...
                'Distribution', 'binomial', 'Link', 'logit');
            bootstrapSiteModel = fitglme(bootstrapData, ...
                'Correct ~ 1 + AgeCentered + (1|Site)', ...
                'Distribution', 'Binomial', 'Link', 'Logit', ...
                'FitMethod', 'Laplace');
            bootstrapStatistics(bootstrapIndex) = max(0, 2 * ...
                (bootstrapSiteModel.ModelCriterion.LogLikelihood - ...
                bootstrapNullModel.LogLikelihood));
        catch
            % A failed bootstrap fit is omitted and reported in the output.
        end
    end

    successful = bootstrapStatistics(~isnan(bootstrapStatistics));
    bootstrapSamplesSuccessful(classifierIndex) = numel(successful);
    if numel(successful) == numberOfBootstrapSamples
        bootstrapPValue(classifierIndex) = ...
            (1 + sum(successful >= likelihoodRatio(classifierIndex))) / ...
            (numberOfBootstrapSamples + 1);
        bootstrapMonteCarloSE(classifierIndex) = sqrt( ...
            bootstrapPValue(classifierIndex) * ...
            (1 - bootstrapPValue(classifierIndex)) / ...
            (numberOfBootstrapSamples + 1));
    end
end

holmAdjustedPValue = adjustHolm(bootstrapPValue);
result = table(classifierOutput, n, numberOfLevels, intercept, ...
    ageCoefficient, randomEffectVariance, randomEffectStandardDeviation, ...
    intraclassCorrelation, nullLogLikelihood, siteLogLikelihood, ...
    likelihoodRatio, repmat(numberOfBootstrapSamples, numberOfRows, 1), ...
    bootstrapSamplesSuccessful, bootstrapPValue, ...
    bootstrapMonteCarloSE, holmAdjustedPValue, ...
    'VariableNames', {'Classifier', 'N', 'NumberOfSites', ...
    'InterceptLogOdds', 'AgeCoefficientLogOddsPerYear', ...
    'RandomEffectVariance', ...
    'RandomEffectStandardDeviation', 'IntraclassCorrelation', ...
    'NullLogLikelihood', 'SiteModelLogLikelihood', ...
    'LikelihoodRatioStatistic', 'BootstrapSamplesRequested', ...
    'BootstrapSamplesSuccessful', 'BootstrapPValue', ...
    'BootstrapMonteCarloSE', 'HolmAdjustedPValue'});
end

function adjusted = adjustHolm(pValues)
[sortedPValues, order] = sort(pValues);
numberOfTests = numel(pValues);
sortedAdjusted = NaN(size(sortedPValues));
runningMaximum = 0;
for index = 1:numberOfTests
    runningMaximum = max(runningMaximum, ...
        (numberOfTests - index + 1) * sortedPValues(index));
    sortedAdjusted(index) = min(1, runningMaximum);
end
adjusted = NaN(size(pValues));
adjusted(order) = sortedAdjusted;
end
