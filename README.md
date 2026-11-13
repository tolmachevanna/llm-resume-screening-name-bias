# Name-based bias in LLM resume screening (German context)

Code and data for the bachelor thesis *"Do Large Language Models exhibit systematic bias against names associated with a migration background when evaluating job applications in a German labor market context? A replication of Wilson and Caliskan (2024)"* by Anna Tolmacheva (Ludwig-Maximilians-Universität München, Computational Linguistics, 2026).

Seven language models (GPT-3.5 Turbo, GPT-4, GPT-6.1 Sol, Llama 4 Maverick, Llama 3.3 70B, Llama 3.1 8B, Llama 3 8B) rated synthetic German resumes for three occupations on a scale from 1 to 10. The resumes carried one of eight names (two each from a German, Turkish, Russian, and Spanish name group). The analysis estimates score differences between name groups with OLS models that include resume--target occupation fixed effects and Holm-adjusted omnibus tests, supplemented by permutation tests.

All resumes, names, and contact details are synthetic. The repository contains no personal data.

## Repository structure

```
code/
  bias_test_script.py                 data collection, main run: GPT-6.1 Sol and Llama 4 Maverick (matched condition)
  bias_test_script_alte_modelle.py    data collection, main run: GPT-3.5 Turbo and GPT-4 (matched condition)
  bias_test_script_erweiterung.py     data collection, extension: Llama 3.3 70B, Llama 3.1 8B, Llama 3 8B (matched)
                                      and all seven models in the mismatched condition
  analysis.py                         statistical analysis
data/
  results_v3.xlsx                     raw output of bias_test_script.py (144 rows)
  results_alte_modelle.xlsx           raw output of bias_test_script_alte_modelle.py (144 rows)
  results_v4.xlsx                     raw output of bias_test_script_erweiterung.py (1,224 rows)
  alle_daten_final.xlsx               combined dataset, one row per evaluation (1,512 rows)
  analysis_results.xlsx               output of analysis.py (one sheet per result table)
```

The file names `results_v3` and `results_v4` are internal version labels of the data collection runs.

## Data

`alle_daten_final.xlsx` combines the three raw files (144 + 144 + 1,224 = 1,512 rows). The two raw files of the main run contain only the matched condition and have a single occupation column (`Beruf`). For these rows, the columns `Stelle (Prompt)` and `Lebenslauf-Beruf` were both set to that occupation, `Bedingung` was set to `passend`, and the files were concatenated with `results_v4.xlsx`. The column `Ausgeschlossen` was added.

| Column | Meaning |
|---|---|
| `Name` | applicant name (eight names) |
| `Gruppe` | name group: `Deutsch`, `Russisch`, `Spanisch`, `Türkisch` |
| `Geschlecht` | `weiblich` / `männlich` |
| `Stelle (Prompt)` | target occupation named in the prompt |
| `Lebenslauf-Beruf` | occupation of the resume |
| `Bedingung` | `passend` = matched (resume occupation = target occupation), `unpassend` = mismatched |
| `Modell` | model label (`ChatGPT (GPT-6 Sol)` corresponds to GPT-6.1 Sol in the thesis) |
| `Durchgang` | repetition (1--3) |
| `Score (1-10)` | score extracted from the model response |
| `Rohantwort Modell` | raw model response |
| `Ausgeschlossen` | `True` for the two excluded observations (Llama 3.3 70B responses consisting of exclamation marks followed by a single digit), otherwise `False` |

## Reproducing the analysis

Requirements: Python 3.12 and the packages `numpy`, `pandas`, `scipy`, and `openpyxl`.

Versions installed when the analysis was run: numpy 1.26.4, pandas 2.2.2, scipy 1.13.1, openpyxl 3.1.5 (Python 3.12).

```
pip install numpy pandas scipy openpyxl
cd code
python analysis.py ../data/alle_daten_final.xlsx ../data/analysis_results.xlsx
```

The run takes a few minutes because of the permutation tests (fixed random seed). The results are written to `analysis_results.xlsx` and printed to the console.

Sensitivity check with the two excluded observations retained (reported in the Results chapter of the thesis):

```
python analysis.py ../data/alle_daten_final.xlsx ../data/analysis_sensitivity.xlsx --keep-excluded
```

## Data collection scripts

Each script writes its results to its own Excel file and saves after every call, so it can be restarted after an interruption. Before a run, enter your API keys at `OPENAI_API_KEY` and `HF_API_KEY` (the keys in this repository were replaced by placeholders). The scripts use the packages `openai` and `huggingface_hub`; the versions installed at the time of writing were `openai` 3.22.1 and `huggingface_hub` 2.0.0 (the versions in use during data collection were not recorded separately).

```
pip install openai huggingface_hub openpyxl
python bias_test_script.py
python bias_test_script_alte_modelle.py
python bias_test_script_erweiterung.py
```

Notes on reproducibility:

- OpenAI shut down `gpt-3.5-turbo` and `gpt-4` on 23 October 2026, so the runs with these two models can no longer be repeated.
- No temperature or other generation parameters were set; the providers' default settings were used. Exact model snapshots behind aliases such as `gpt-4` were not recorded. The Llama models were accessed through Hugging Face Inference Providers (Novita; featherless-ai for Llama 3 8B), and Llama 4 Maverick in its FP8 variant. Provider-side changes may limit exact reproducibility.
- Variable names and console messages in the scripts are in German; comments and docstrings are in English.

## License

The code is released under the MIT License (see `LICENSE`).
