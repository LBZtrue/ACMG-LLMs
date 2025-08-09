import os
import re
import json
from typing import List, Dict, Any

# This code is used for formatting and standardizing the JSON responses output by LLMs

class JSONStandardizer:
    """Standardizes JSON output from different models (Gemini-2.0-flash, GPT-4o, Qwen-Long)"""

    @staticmethod
    def remove_markdown_codeblock_markers(text: str) -> str:
        """
        Remove Markdown code block markers (```language) while preserving content.

        Args:
            text (str): Text containing Markdown code blocks

        Returns:
            str: Text with code block markers removed
        """
        pattern = re.compile(r'```(?:\w+)?\n(.*?)```', re.DOTALL)
        
        def replace_codeblock(match):
            return match.group(1)
        
        return pattern.sub(replace_codeblock, text)

    @staticmethod
    def extract_json_from_text(text: str) -> List[Dict[str, Any]]:
        """
        Extract all JSON blocks from text and intelligently identify structure.

        Args:
            text (str): Input text containing JSON data

        Returns:
            List[Dict[str, Any]]: List of extracted JSON objects
        """
        json_pattern = re.compile(r'```json(.*?)```', re.DOTALL)
        matches = json_pattern.finditer(text)
        
        variants = []
        for match in matches:
            try:
                json_data = json.loads(match.group(1).strip())
                
                # Handle different JSON structures
                if isinstance(json_data, dict):
                    if 'functional_evidence_assessment' in json_data:
                        variants.extend(json_data['functional_evidence_assessment'])
                    elif 'variant_id' in json_data:
                        variants.append(json_data)
                elif isinstance(json_data, list) and all(isinstance(item, dict) and 'variant_id' in item for item in json_data):
                    variants.extend(json_data)
                    
            except json.JSONDecodeError as e:
                print(f"JSON parsing error: {str(e)}")
        
        return variants

    @staticmethod
    def standardize_variant_id(variant_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Standardize variant_id field, ensuring position is integer type.

        Args:
            variant_data (Dict[str, Any]): Variant data dictionary

        Returns:
            Dict[str, Any]: Standardized variant data
        """
        if 'variant_id' in variant_data:
            protein_change = variant_data['variant_id'].get('Protein_Change', {})
            if 'position' in protein_change:
                try:
                    protein_change['position'] = int(protein_change['position'])
                except (ValueError, TypeError):
                    print(f"Warning: Could not convert position to integer in variant_id: {protein_change['position']}")
        return variant_data

    @staticmethod
    def standardize_assessment_steps(variant_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Standardize assessment steps, adding missing steps.

        Args:
            variant_data (Dict[str, Any]): Variant data dictionary

        Returns:
            Dict[str, Any]: Variant data with standardized assessment steps
        """
        required_steps = [
            "Step 1: Define the disease mechanism",
            "Step 2: Evaluate applicability of general classes of assay used in the field",
            "Step 3: Evaluate validity of specific instances of assays",
            "Step 4: Apply evidence to individual variant interpretation"
        ]
        
        if 'assessment_steps' in variant_data and isinstance(variant_data['assessment_steps'], list):
            existing_steps = {step.get('step_name', ''): step for step in variant_data['assessment_steps']}
            standardized_steps = []
            
            for step_name in required_steps:
                existing_step = existing_steps.get(step_name)
                if existing_step:
                    standardized_steps.append(existing_step)
                else:
                    merged_step = JSONStandardizer.merge_substeps(existing_steps, step_name)
                    if merged_step:
                        standardized_steps.append(merged_step)
                    else:
                        standardized_steps.append({
                            "step_name": step_name,
                            "extracted_paper_info": "Not evaluated",
                            "judgment": "Not evaluated",
                            "reasoning": "No information provided in the paper"
                        })
            
            variant_data['assessment_steps'] = standardized_steps
        
        return variant_data

    @staticmethod
    def merge_substeps(existing_steps: Dict[str, Any], target_step: str) -> Dict[str, Any]:
        """
        Merge substeps into a main step.

        Args:
            existing_steps (Dict[str, Any]): Existing steps dictionary
            target_step (str): Target step name to merge

        Returns:
            Dict[str, Any]: Merged step dictionary or None if no substeps found
        """
        substep_mapping = {
            "Step 3: Evaluate validity of specific instances of assays": [
                "Step 3a: Basic Controls and Replicates",
                "Step 3b: Appropriate Comparators",
                "Step 3c: Variant Controls"
            ],
            "Step 4: Apply evidence to individual variant interpretation": [
                "Step 4a: OddsPath Calculation",
                "Step 4b: No OddsPath Calculation"
            ]
        }
        
        substeps = substep_mapping.get(target_step, [])
        relevant_substeps = [existing_steps.get(ss) for ss in substeps if ss in existing_steps]
        
        if not relevant_substeps:
            return None
        
        merged_info = "\n\n".join([
            f"{step.get('step_name', '')}: {step.get('extracted_paper_info', 'Not evaluated')}" 
            for step in relevant_substeps
        ])
        
        merged_reasoning = "\n\n".join([
            f"{step.get('step_name', '')}: {step.get('reasoning', 'No reasoning provided')}" 
            for step in relevant_substeps
        ])
        
        all_yes = all(step.get('judgment', 'No') == 'Yes' for step in relevant_substeps)
        
        return {
            "step_name": target_step,
            "extracted_paper_info": merged_info,
            "judgment": "Yes" if all_yes else "Partial",
            "reasoning": merged_reasoning
        }

    @staticmethod
    def standardize_final_evidence(variant_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Standardize final evidence strength field.

        Args:
            variant_data (Dict[str, Any]): Variant data dictionary

        Returns:
            Dict[str, Any]: Variant data with standardized evidence strength
        """
        if 'final_evidence_strength' in variant_data:
            evidence = variant_data['final_evidence_strength']
            if 'type' in evidence:
                evidence['type'] = evidence['type'].lower().capitalize()
        return variant_data

    @staticmethod
    def wrap_in_standard_structure(variants: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Wrap standardized variants into standard structure.

        Args:
            variants (List[Dict[str, Any]]): List of standardized variants

        Returns:
            Dict[str, Any]: Wrapped JSON structure
        """
        return {
            "functional_evidence_assessment": variants
        }

    @staticmethod
    def process_model_output(input_path: str, output_path: str, model_type: str) -> None:
        """
        Process model output files based on model type.

        Args:
            input_path (str): Input directory path
            output_path (str): Output directory path
            model_type (str): Model type ('gemini', 'gpt4o', or 'qwen')
        """
        os.makedirs(output_path, exist_ok=True)
        
        for filename in os.listdir(input_path):
            if filename.endswith(('.json', '.txt')):
                input_file = os.path.join(input_path, filename)
                output_file = os.path.join(output_path, filename if filename.endswith('.json') else filename.replace('.txt', '.json'))
                
                try:
                    with open(input_file, 'r', encoding='utf-8') as file:
                        content = file.read()
                    
                    if model_type in ['gemini', 'gpt4o']:
                        # For Gemini and GPT-4o: Remove Markdown and save as-is
                        processed_content = JSONStandardizer.remove_markdown_codeblock_markers(content)
                        with open(output_file, 'w', encoding='utf-8') as file:
                            file.write(processed_content)
                    else:  # Qwen-Long
                        # For Qwen: Full standardization process
                        variants = JSONStandardizer.extract_json_from_text(content)
                        standardized_variants = []
                        
                        for variant in variants:
                            variant = JSONStandardizer.standardize_variant_id(variant)
                            variant = JSONStandardizer.standardize_assessment_steps(variant)
                            variant = JSONStandardizer.standardize_final_evidence(variant)
                            standardized_variants.append(variant)
                        
                        standardized_results = JSONStandardizer.wrap_in_standard_structure(standardized_variants)
                        
                        with open(output_file, 'w', encoding='utf-8') as outfile:
                            json.dump(standardized_results, outfile, indent=2, ensure_ascii=False)
                    
                    print(f"Successfully processed file: {filename}")
                
                except FileNotFoundError:
                    print(f"Error: File not found: {input_file}")
                except Exception as e:
                    print(f"Error processing file {filename}: {str(e)}")

if __name__ == "__main__":
    # Configuration for input and output paths
    base_path = "/Users/liuchenbin/Library/CloudStorage/OneDrive-个人/VsCode/ps4_llm/ps3_test/prompt_v3_test"
    model_configs = {
        'gemini': {
            'input': os.path.join(base_path, 'gemini_output'),
            'output': os.path.join(base_path, 'gemini_output/standardized_json')
        },
        'gpt4o': {
            'input': os.path.join(base_path, 'gpt_4o_output'),
            'output': os.path.join(base_path, 'gpt_4o_output/standardized_json')
        },
        'qwen': {
            'input': os.path.join(base_path, 'qwen_output'),
            'output': os.path.join(base_path, 'qwen_output/standardized_json')
        }
    }

    # Process files for each model
    for model_type, paths in model_configs.items():
        print(f"\nProcessing {model_type.upper()} output...")
        JSONStandardizer.process_model_output(paths['input'], paths['output'], model_type)
    print("\nAll model outputs processed successfully!")