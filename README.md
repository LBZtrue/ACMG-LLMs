# PS3-BS3-LLMs-Evaluation

<img width="12219" height="5363" alt="02研究架构主图" src="https://github.com/user-attachments/assets/d25f00ad-4473-4281-8e38-1ce9232980d5" />

## 1、Directory Structure Description
**01.Gold-Standard-Json:** Directory containing expert-curated intermediate information and final PS3/BS3 classifications, serving as the ground-truth benchmark.

**02.Prompts:** Directory storing the various prompts used to test the large language models.

**03.Source-Code:**

**04.LLMs-Generated-Responses:** Directory holding the models’ initial responses under different prompts (Raw QA Outputs, including reasoning traces) and their corresponding structured JSON outputs (Normalized JSON Responses).

**05.Evaluation-Results:**

## 2、Tested LLM Inventory
| Environment | Model Name          | Parameters | RAG Mode      |
|:------------|:--------------------|:-----------|:--------------|
| Online ↓    | Doubao-1.5-Pro      | —          | noRAG   |
|             | DeepSeek-v3         | 671 B      | noRAG   |
|             | Gemini-2.0-flash    | —          | noRAG   |
|             | GPT-4o              | 1.8 T      | noRAG   |
|             | Grok3               | 2.7 T      | noRAG   |
|             | Kimi-v1             | —          | noRAG   |
|             | Qwen-Long           | 235 B      | noRAG   |
| Local ↓     | Llama3.1:70b        | 70 B       | RAG   |
|             | Llama3.1:8b         | 8 B        | noRAG / RAG   |
|             | Gemma3:27b          | 27 B       | RAG   |
|             | Gemma3:12b          | 12 B       | noRAG / RAG   |
|             | Gemma3:4b           | 4 B        | RAG   |
|             | Gemma3:1b           | 1 B        | RAG   |
|             | Qwen3:32b           | 32 B       | RAG   |
|             | Qwen3:30b-a3b       | 30 B-A3B   | RAG   |
|             | Qwen3:14b           | 14 B       | noRAG / RAG   |
|             | Qwen3:8b            | 8 B        | RAG   |
|             | Qwen3:4b            | 4 B        | RAG   |
|             | Qwen3:1.7b          | 1.7 B      | RAG   |
|             | Qwen3:0.6b          | 0.6 B      | RAG   |
|             | Mistral:7b          | 7 B        | noRAG / RAG   |
|             | DeepSeek-r1:70b     | 70 B       | RAG   |
|             | DeepSeek-r1:32b     | 32 B       | RAG   |
|             | DeepSeek-r1:14b     | 14 B       | noRAG / RAG   |
|             | DeepSeek-r1:7b      | 7 B        | RAG   |
|             | DeepSeek-r1:1.5b    | 1.5 B      | RAG   |

## 3、Additional Remarks

