# ACMG-PS3/BS3-LLMs-Evaluation
## 1、Comprehensive Framework for LLMs Evaluation
<img width="12219" height="5363" alt="02研究架构主图" src="https://github.com/user-attachments/assets/d25f00ad-4473-4281-8e38-1ce9232980d5" />


## 2、Directory Structure Description

**01.Gold-Standard-Json:** Directory containing expert-curated intermediate information and final PS3/BS3 classifications, serving as the ground-truth benchmark.

**02.Prompts:** Directory storing the various prompts used to test the large language models.

**03.Source-Code:** Code repository for JSON-formatting utilities and evaluation scripts.

**04.LLMs-Generated-Responses:** Directory holding the models’ initial responses under different prompts (Raw QA Outputs, including reasoning traces) and their corresponding structured JSON outputs (Normalized JSON Responses).

**05.Evaluation-Results:** Dataset repository housing all evaluation results (The intermediate data, final results, and visualization of outcomes are detailed in the relevant paper and its supplementary materials).


## 3、Tested LLM Inventory
### (1) Local Deployment LLMs Test Parameter Inventory

| Model Name  | 235B-A22B | 70B | 32B | 30B-A3B | 27B | 14B | 12B | 8B | 7B | 4B | 1.7B | 1.5B | 1B | 0.6B |
| :---------- | :-------- | :-- | :-- | :------ | :-- | :-- | :-- | :- | :- | :- | :--- | :--- | :- | :--- |
| Llama3.1    |           | ✓   |     |         |     |     |     | ✓  |    |    |      |      |    |      |
| Gemma3      |           |     |     |         | ✓   |     | ✓   |    |    | ✓  |      |      | ✓  |      |
| Qwen3       | ✓         |     | ✓   | ✓       |     | ✓   |     | ✓  |    | ✓  | ✓    |      |    | ✓    |
| Mistral     |           |     |     |         |     |     |     |    | ✓  |    |      |      |    |      |
| DeepSeek-r1 |           | ✓   | ✓   |         |     | ✓   |     |    | ✓  |    |      | ✓    |    |      |


### (2) Comprehensive Test Inventory of All LLMs

| Environment | Model Name | Parameters | RAG Mode |
|:------------|:-----------|:-----------|:---------|
| **Online ↓** |
|             | Doubao-1.5-Pro | — | noRAG |
|             | DeepSeek-v3 | 671B | noRAG |
|             | Gemini-2.0-flash | — | noRAG |
|             | GPT-4o | 1.8T | noRAG |
|             | Grok3 | 2.7T | noRAG |
|             | Kimi-v1 | — | noRAG |
|             | Qwen-Long | 235B | noRAG |
| **Local ↓** |
|             | Llama3.1 | 70B / 8B | RAG / noRAG·RAG |
|             | Gemma3 | 27B / 12B / 4B / 1B | RAG / noRAG·RAG / RAG / RAG |
|             | Qwen3 | 235B-A22B / 32B / 30B-A3B / 14B / 8B / 4B / 1.7B / 0.6B | RAG / RAG / RAG / noRAG·RAG / RAG / RAG / RAG / RAG |
|             | Mistral | 7B | noRAG·RAG |
|             | DeepSeek-r1 | 70B / 32B / 14B / 7B / 1.5B | RAG / RAG / noRAG·RAG / RAG / RAG |


## 4、Additional Remarks
### (1) PS3/BS3 Evidence Assessment Process and Visualization of JSON Information Structure

<img width="6444" height="3663" alt="00综合图2" src="https://github.com/user-attachments/assets/5495aabd-7bcc-4b7e-80b3-25615ef6be65" />

**Left Figure:** PS3/BS3 Evaluation Based on ACMG/AMP–ClinGen SVI Criteria.

**Right Figure:** A circular tree structure diagram is utilized, with the root node positioned at the center, extending outward to the first layer of nodes, which primarily include information pertaining to genes, diseases, variants, and experimental methods. This structure progressively expands to the outermost leaf nodes, which store detailed information. Such a hierarchical tree structure enables efficient data storage, retrieval, and analysis.

### (2) Under revision ...
