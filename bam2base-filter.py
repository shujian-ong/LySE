import pysam
import pandas as pd
from collections import defaultdict

def analyze_bam_quality(bam_file, reference_fasta, output_file, min_qual=30):
    """
    Analyze BAM file nucleotide composition with quality score filtering.
    
    Parameters:
    bam_file (str): Path to input BAM file
    reference_fasta (str): Path to reference genome FASTA file
    output_file (str): Path for output CSV file
    min_qual (int): Minimum quality score threshold (default: 30)
    """
    # Open BAM and reference files
    bam = pysam.AlignmentFile(bam_file, "rb")
    reference = pysam.FastaFile(reference_fasta)
    
    # Store data for each position
    position_data = defaultdict(lambda: {
        'A': 0, 'T': 0, 'C': 0, 'G': 0, 'N': 0,
        'total_reads': 0,
        'qual_scores': [],
        'ref_base': '',
        'filtered_reads': 0  # Count of reads passing quality filter
    })
    
    # Process each read in the BAM file
    for read in bam:
        if read.is_unmapped:
            continue
            
        ref_name = read.reference_name
        qualities = read.query_qualities
        
        for read_pos, ref_pos in read.get_aligned_pairs(matches_only=True):
            if ref_pos is None:
                continue
                
            qual = qualities[read_pos]
            base = read.query_sequence[read_pos].upper()
            
            position = (ref_name, ref_pos)
            position_data[position]['total_reads'] += 1
            
            # Only count bases that pass quality threshold
            if qual >= min_qual:
                position_data[position][base] += 1
                position_data[position]['filtered_reads'] += 1
                position_data[position]['qual_scores'].append(qual)
    
    # Process results and create DataFrame
    results = []
    for (ref_name, pos), data in position_data.items():
        # Get reference base
        ref_base = reference.fetch(ref_name, pos, pos + 1).upper()
        
        # Calculate percentages and quality metrics
        total_filtered = data['filtered_reads']
        
        row = {
            'reference': ref_name,
            'position': pos + 1,  # 1-based position
            'ref_base': ref_base,
            'total_coverage': data['total_reads'],
            'high_quality_coverage': total_filtered,
            'mean_quality': sum(data['qual_scores']) / len(data['qual_scores']) if data['qual_scores'] else 0,
            'A_pct': (data['A'] / total_filtered * 100) if total_filtered > 0 else 0,
            'T_pct': (data['T'] / total_filtered * 100) if total_filtered > 0 else 0,
            'C_pct': (data['C'] / total_filtered * 100) if total_filtered > 0 else 0,
            'G_pct': (data['G'] / total_filtered * 100) if total_filtered > 0 else 0,
            'N_pct': (data['N'] / total_filtered * 100) if total_filtered > 0 else 0,
            'A_count': data['A'],
            'T_count': data['T'],
            'C_count': data['C'],
            'G_count': data['G'],
            'N_count': data['N']
        }
        
        # Add note if no high-quality reads
        if total_filtered == 0:
            row['note'] = f"No reads met minimum quality threshold (Q{min_qual})"
        else:
            row['note'] = ''
            
        results.append(row)
    
    # Create and sort DataFrame
    df = pd.DataFrame(results)
    df = df.sort_values(['reference', 'position'])
    
    # Save to CSV
    df.to_csv(output_file, index=False)
    
    # Print summary statistics
    print("\nAnalysis Summary:")
    print(f"Total positions analyzed: {len(df)}")
    print(f"Positions with no high-quality reads: {len(df[df['high_quality_coverage'] == 0])}")
    print(f"Average high-quality coverage: {df['high_quality_coverage'].mean():.2f}")
    
    return df

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Analyze BAM file nucleotide composition with quality filtering')
    parser.add_argument('--bam', required=True, help='Input BAM file')
    parser.add_argument('--reference', required=True, help='Reference genome FASTA file')
    parser.add_argument('--output', required=True, help='Output CSV file')
    parser.add_argument('--min-qual', type=int, default=30, help='Minimum quality score (default: 30)')
    
    args = parser.parse_args()
    
    analyze_bam_quality(args.bam, args.reference, args.output, args.min_qual)