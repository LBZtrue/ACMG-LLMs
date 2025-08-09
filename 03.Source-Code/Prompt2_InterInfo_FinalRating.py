import json
import pandas as pd
import os

# Obtain the final rating result based on the intermediate information extracted by LLMs

def determine_strength_by_oddpath(odds_path, is_perfect_binary=None):
    """Determine evidence strength based on OddsPath value and conditions"""
    strength = "Supporting"
    if (odds_path < 0.0029) or (odds_path > 350):
        strength = "Very Strong"
    elif (0.0029 <= odds_path < 0.053) or (18.7 < odds_path <= 350):
        strength = "Strong"
    elif (0.053 <= odds_path < 0.23) or (4.3 < odds_path <= 18.7):
        strength = "Moderate"
    return strength

def determine_evidence_strength(data):
    """Determine evidence strength for gene variant-disease association based on multi-level conditions"""
    # First level: Check if experimental method is approved
    if not evaluate_assay_validity_approved(data):
        return "No PS3/BS3"

    # Second level: Check if experiment includes valid controls and replicates
    if not evaluate_assay_validity_control(data):
        return "No PS3/BS3"

    # Third level: Check if experiment includes known pathogenic/benign variants
    if not evaluate_assay_contains_known_variants(data):
        return "Supporting"

    # Fourth level: Check if OddsPath can be calculated
    can_calculate_oddpath, odds_path, is_perfect_binary = calculate_oddpath(data)
    if not can_calculate_oddpath:
        # Cannot calculate OddsPath, count total pathogenic/benign variants
        pathogenic_count, benign_count = count_pathogenic_benign_variants(data)
        total_count = pathogenic_count + benign_count
        if total_count > 10:
            return "Moderate"
        else:
            return "Supporting"
    else:
        # Determine evidence strength based on OddsPath value and conditions
        return determine_strength_by_oddpath(odds_path, is_perfect_binary)

def evaluate_assay_contains_known_variants(data):
    """Check if experiment includes known pathogenic/benign variants"""
    try:
        for assay in data["Experiment Method"]:
            # Process Validation controls P/LP field
            pathogenic_field = assay.get("Validation controls P/LP Alexandria")
            if isinstance(pathogenic_field, dict):
                has_pathogenic = pathogenic_field.get("Validation controls P/LP") == "Yes"
            elif pathogenic_field is None:
                has_pathogenic = False
            else:
                has_pathogenic = pathogenic_field == "Yes"

            # Process Validation controls B/LB field
            benign_field = assay.get("Validation controls B/LB")
            if isinstance(benign_field, dict):
                has_benign = benign_field.get("Validation controls B/LB") == "Yes"
            elif benign_field is None:
                has_benign = False
            else:
                has_benign = benign_field == "Yes"

            if has_pathogenic or has_benign:
                return True
        return False
    except (KeyError, TypeError):
        return False

def calculate_oddpath(data):
    """Calculate oddpath value and determine condition (perfect binary or allows one indeterminate reading)"""
    pathogenic_count = 0
    benign_count = 0
    indeterminate_count = 0

    try:
        for assay in data["Experiment Method"]:
            # Get pathogenic variant count from Validation controls P/LP
            p_field = assay.get("Validation controls P/LP")
            p_counts = p_field.get("Counts", 0) if isinstance(p_field, dict) else 0
            pathogenic_count += int(p_counts) if str(p_counts).isdigit() else 0

            # Get benign variant count from Validation controls B/LB
            b_field = assay.get("Validation controls B/LB")
            b_counts = b_field.get("Counts", 0) if isinstance(b_field, dict) else 0
            benign_count += int(b_counts) if str(b_counts).isdigit() else 0

            # Count indeterminate conclusions
            readout = assay.get("Readout description")
            if isinstance(readout, list):
                for result in readout:
                    if isinstance(result, dict) and result.get("Conclusion") == "Indeterminate":
                        indeterminate_count += 1

        # Calculate prior probability P1
        total_variants = pathogenic_count + benign_count
        if total_variants > 0:
            p1 = (pathogenic_count + 1) / (total_variants + 2)
            p2 = (pathogenic_count + 1) / (total_variants + 2)
            odds_path = (p2 * (1 - p1)) / ((1 - p2) * p1)
        else:
            return False, 0.0, True

        if indeterminate_count == 0:
            return True, odds_path, True
        elif indeterminate_count == 1:
            return True, odds_path, False
        else:
            return False, 0.0, True
    except ZeroDivisionError:
        return False, 0.0, True
    except (ValueError, TypeError):
        return False, 0.0, True

def count_pathogenic_benign_variants(data):
    """Count the number of pathogenic and benign variants in the experiment"""
    pathogenic_count = 0
    benign_count = 0

    try:
        for assay in data["Experiment Method"]:
            readout = assay.get("Readout description")
            if not readout:
                continue

            if isinstance(readout, list):
                for result in readout:
                    if isinstance(result, dict):
                        conclusion = result.get("Conclusion")
                        if conclusion == "Abnormal":
                            pathogenic_count += 1
                        elif conclusion == "Normal":
                            benign_count += 1
            elif isinstance(readout, str):
                # String format cannot distinguish different variants, skip counting
                pass
    except (KeyError, TypeError):
        pass

    return pathogenic_count, benign_count

def evaluate_assay_validity_approved(data):
    """Check if experimental method is Approved (at least one method marked as Yes)"""
    try:
        for assay in data.get("Experiment Method", []):
            # Process Approved assay field
            approved_field = assay.get("Approved assay")
            if approved_field is None:
                continue

            if isinstance(approved_field, dict):
                approved_value = approved_field.get("Approved assay", "No")
            else:
                approved_value = approved_field

            if approved_value == "Yes":
                return True
        return False
    except (KeyError, TypeError):
        return False

def evaluate_assay_validity_control(data):
    """Evaluate if experiment includes valid controls and replicates"""
    try:
        for assay in data["Experiment Method"]:
            # Process Basic positive control
            basic_pos_control = assay.get("Basic positive control")
            if basic_pos_control is None:
                has_basic_pos = False
            elif isinstance(basic_pos_control, dict):
                has_basic_pos = basic_pos_control.get("Basic positive control") == "Yes"
            else:
                has_basic_pos = basic_pos_control == "Yes"

            # Process Basic negative control
            basic_neg_control = assay.get("Basic negative control")
            if basic_neg_control is None:
                has_basic_neg = False
            elif isinstance(basic_neg_control, dict):
                has_basic_neg = basic_neg_control.get("Basic negative control") == "Yes"
            else:
                has_basic_neg = basic_neg_control == "Yes"

            # Check biological replicates
            bio_replicates = assay.get("Biological replicates")
            if bio_replicates is None:
                has_bio_reps = False
            elif isinstance(bio_replicates, dict):
                has_bio_reps = bio_replicates.get("Biological replicates") == "Yes"
            else:
                has_bio_reps = bio_replicates == "Yes"

            # Check technical replicates
            tech_replicates = assay.get("Technical replicates")
            if tech_replicates is None:
                has_tech_reps = False
            elif isinstance(tech_replicates, dict):
                has_tech_reps = tech_replicates.get("Technical replicates") == "Yes"
            else:
                has_tech_reps = tech_replicates == "Yes"

            # Check basic controls (at least one positive or negative control)
            has_basic_control = has_basic_pos or has_basic_neg
            # Check replicates (at least one biological or technical replicate)
            has_replicates = has_bio_reps or has_tech_reps

            if has_basic_control and has_replicates:
                return True
        return False
    except (KeyError, TypeError):
        return False

def read_json_file(file_path):
    """Read JSON file and return data"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except (FileNotFoundError, json.JSONDecodeError, Exception) as e:
        print(f"Error: {e}")
        return None

def analyze_variants_evidence(data):
    """Analyze pathogenicity evidence strength for each variant"""
    variants = {}
    # Extract all variant information
    for gene_info in data.get("Variants Include", []):
        for var in gene_info.get("variants", []):
            hgvs = var.get("HGVS")
            if hgvs:
                variants[hgvs] = {
                    "HGVS": hgvs,
                    "Protein Change": var.get("Protein Change"),
                    "Descriptions": [],
                    "Conclusions": [],
                    "is_pathogenic": None,
                    "evidence_strength": "No PS3/BS3"
                }

    # Collect conclusions for each variant across different experiments
    for assay in data.get("Experiment Method", []):
        assay_name = assay.get("Assay Method", "Unknown Assay")
        readout = assay.get("Readout description")

        # Case 1: Readout description is a list
        if isinstance(readout, list):
            for result in readout:
                if not isinstance(result, dict):
                    continue

                variant_hgvs = result.get("Variant")
                conclusion = result.get("Conclusion")
                if variant_hgvs in variants and conclusion:
                    variants[variant_hgvs]["Conclusions"].append({
                        "assay": assay_name,
                        "conclusion": conclusion,
                        "molecular_effect": result.get("Molecular Effect"),
                        "description": result.get("Result Description")
                    })

        # Case 2: Has Readout details
        elif "Readout details" in assay:
            readout_details = assay.get("Readout details", {})
            for gene, gene_data in readout_details.items():
                for var_desc, description in gene_data.items():
                    # Find matching variant
                    for hgvs, variant in variants.items():
                        if var_desc == hgvs or (variant["Protein Change"] and
                                                var_desc == f"{variant['Protein Change'].get('ref', '')}"
                                                            f"{variant['Protein Change'].get('position', '')}"
                                                            f"{variant['Protein Change'].get('alt', '')}"):
                            variants[hgvs]["Conclusions"].append({
                                "assay": assay_name,
                                "conclusion": "Abnormal" if "Increased" in description or "Reduced" in description else "Unknown",
                                "molecular_effect": description,
                                "description": description
                            })

        # Case 3: Readout description is a string
        elif isinstance(readout, str):
            # Attempt to extract conclusions from description
            for hgvs in variants:
                if hgvs in readout:
                    variants[hgvs]["Conclusions"].append({
                        "assay": assay_name,
                        "conclusion": "Abnormal" if "Increased" in readout or "Reduced" in readout else "Normal",
                        "molecular_effect": readout,
                        "description": readout
                    })

    # Evaluate pathogenicity and evidence strength for each variant
    for hgvs, var_data in variants.items():
        conclusions = [concl["conclusion"] for concl in var_data["Conclusions"]]

        # Determine pathogenicity
        if all(concl == "N.D." for concl in conclusions):
            var_data["is_pathogenic"] = None
        else:
            var_data["is_pathogenic"] = any(concl == "Abnormal" for concl in conclusions)

        # Create subset data for this variant - use original exact match method
        variant_data = {
            "Article Info": data.get("Article Info", {}),
            "Experiment Method": []
        }

        # Select experiments using original exact match method
        for assay in data.get("Experiment Method", []):
            readout = assay.get("Readout description")
            if not readout:
                continue

            found = False
            # Process list format Readout
            if isinstance(readout, list):
                for result in readout:
                    if isinstance(result, dict) and result.get("Variant") == hgvs:
                        found = True
                        break
            # Process Readout details format
            elif "Readout details" in assay:
                readout_details = assay.get("Readout details", {})
                for gene, gene_data in readout_details.items():
                    if hgvs in gene_data:
                        found = True
                        break
            # Process string format
            elif isinstance(readout, str) and hgvs in readout:
                found = True

            if found:
                variant_data["Experiment Method"].append(assay)

        # Calculate evidence strength using existing logic
        if variant_data["Experiment Method"]:
            strength = determine_evidence_strength(variant_data)
            var_data["evidence_strength"] = strength
        else:
            var_data["evidence_strength"] = "No PS3/BS3"

    return list(variants.values())

def save_variant_analysis_to_csv(data, filename="variant_analysis.csv"):
    """Save variant analysis results to CSV file"""
    variant_analyses = analyze_variants_evidence(data)

    all_assays = set()
    for analysis in variant_analyses:
        for concl in analysis["Conclusions"]:
            all_assays.add(concl["assay"])
    all_assays = sorted(list(all_assays))

    csv_data = []
    for analysis in variant_analyses:
        protein_change = analysis["Protein Change"]
        change_desc = f"{protein_change.get('ref', '?')}{protein_change.get('position', '?')}{protein_change.get('alt', '?')}"

        pathogenic_text = "Unknown" if analysis["is_pathogenic"] is None else (
            "Pathogenic" if analysis["is_pathogenic"] else "Benign")

        row = {
            "Variant": f"{analysis['HGVS']} ({change_desc})",
            "Pathogenicity": pathogenic_text,
            "Evidence Strength": analysis["evidence_strength"]
        }

        assay_data = {concl["assay"]: concl for concl in analysis["Conclusions"]}
        for assay in all_assays:
            if assay in assay_data:
                concl = assay_data[assay]
                row[f"Experiment Conclusion({assay})-conclusion"] = concl["conclusion"]
                row[f"Experiment Conclusion({assay}-molecular_effect"] = concl["molecular_effect"]
                row[f"Experiment Conclusion({assay})-description"] = concl["description"]
            else:
                row[f"Experiment Conclusion({assay})-conclusion"] = "Not Tested"
                row[f"Experiment Conclusion({assay})-molecular_effect"] = ""
                row[f"Experiment Conclusion({assay})-description"] = ""

        csv_data.append(row)

    df = pd.DataFrame(csv_data)
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    print(f"Variant analysis results saved to {filename}")
    return df

def print_variant_analysis(data):
    """Print variant analysis results"""
    variant_analyses = analyze_variants_evidence(data)

    print("Variant Evidence Strength Analysis Results:")
    for analysis in variant_analyses:
        protein_change = analysis["Protein Change"]
        change_desc = f"{protein_change.get('ref', '?')}{protein_change.get('position', '?')}{protein_change.get('alt', '?')}"

        pathogenic_text = "Unknown" if analysis["is_pathogenic"] is None else (
            "Pathogenic" if analysis["is_pathogenic"] else "Benign")

        print(f"\nVariant: {analysis['HGVS']} ({change_desc})")
        print(f"Pathogenicity: {pathogenic_text}")
        print(f"Evidence Strength: {analysis['evidence_strength']}")

        print("Experiment Conclusions:")
        for i, concl in enumerate(analysis["Conclusions"], 1):
            print(f"  {i}. {concl['assay']}: {concl['conclusion']}")
            print(f"     - Molecular Effect: {concl['molecular_effect']}")
            print(f"     - Description: {concl['description']}")


if __name__ == "__main__":
    file_path = "test/11812148_qwen3_01.json"
    json_data = read_json_file(file_path)
    if json_data:
        print_variant_analysis(json_data)
        save_variant_analysis_to_csv(json_data, "variant_analysis.csv")
