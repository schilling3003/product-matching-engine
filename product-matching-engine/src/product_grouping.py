import numpy as np
import pandas as pd
from typing import List, Tuple, Dict, Any

try:
    import networkx as nx
    HAS_NETWORKX = True
except ImportError:
    HAS_NETWORKX = False


def _to_serializable_scalar(value: Any) -> Any:
    """Convert non-scalar values into Arrow-safe scalar representations."""
    if isinstance(value, (dict, list, set, tuple)):
        return str(value)
    if isinstance(value, (np.integer, np.floating)):
        return value.item()
    return value


class UnionFind:
    """Union-Find (Disjoint Set) data structure for efficient connected component finding."""
    
    def __init__(self, n: int):
        self.parent = list(range(n))
        self.rank = [0] * n
    
    def find(self, x: int) -> int:
        """Find the root of x with path compression."""
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]
    
    def union(self, x: int, y: int) -> None:
        """Union the sets containing x and y."""
        px, py = self.find(x), self.find(y)
        if px == py:
            return
        
        # Union by rank
        if self.rank[px] < self.rank[py]:
            self.parent[px] = py
        elif self.rank[px] > self.rank[py]:
            self.parent[py] = px
        else:
            self.parent[py] = px
            self.rank[px] += 1


def find_product_groups(similarity_matrix: np.ndarray, 
                       threshold: float = 80.0,
                       method: str = 'union_find') -> List[List[int]]:
    """
    Find groups of similar products using connected components approach.
    
    Args:
        similarity_matrix: NxN matrix of similarity scores
        threshold: Minimum similarity score to consider two products connected
        method: 'union_find' for custom implementation or 'networkx' for NetworkX
        
    Returns:
        List of groups, where each group is a list of product indices
    """
    n = len(similarity_matrix)
    
    if method == 'networkx':
        if not HAS_NETWORKX:
            print("Warning: NetworkX not installed, falling back to Union-Find method")
            method = 'union_find'
        else:
            return _find_groups_networkx(similarity_matrix, threshold)
    
    if method == 'union_find':
        return _find_groups_union_find(similarity_matrix, threshold)
    
    # Default fallback
    return _find_groups_union_find(similarity_matrix, threshold)


def _find_groups_union_find(similarity_matrix: np.ndarray, threshold: float) -> List[List[int]]:
    """Find groups using Union-Find algorithm."""
    n = len(similarity_matrix)
    uf = UnionFind(n)
    
    # Create unions for similar pairs
    for i in range(n):
        # Only check upper triangle to avoid duplicate work
        for j in range(i + 1, n):
            if similarity_matrix[i][j] >= threshold:
                uf.union(i, j)
    
    # Group by parent
    groups = {}
    for i in range(n):
        root = uf.find(i)
        if root not in groups:
            groups[root] = []
        groups[root].append(i)
    
    # Convert to list and filter single-item groups
    result = [group for group in groups.values() if len(group) > 1]
    return result


def _find_groups_networkx(similarity_matrix: np.ndarray, threshold: float) -> List[List[int]]:
    """Find groups using NetworkX connected components."""
    n = len(similarity_matrix)
    G = nx.Graph()
    G.add_nodes_from(range(n))
    
    # Add edges for similar pairs
    for i in range(n):
        for j in range(i + 1, n):
            if similarity_matrix[i][j] >= threshold:
                G.add_edge(i, j, weight=similarity_matrix[i][j])
    
    # Find connected components
    components = list(nx.connected_components(G))
    
    # Filter single-item groups and convert to lists
    result = [list(component) for component in components if len(component) > 1]
    return result


def analyze_groups(similarity_matrix: np.ndarray,
                  groups: List[List[int]],
                  product_names: List[str] = None,
                  threshold: float = None,
                  conservative: bool = False) -> List[Dict[str, Any]]:
    """
    Analyze product groups to calculate statistics and find representatives.
    
    Args:
        similarity_matrix: NxN matrix of similarity scores
        groups: List of groups (each a list of product indices)
        product_names: Optional list of product names for display
        
    Returns:
        List of group analysis dictionaries
    """
    analyses = []

    if conservative and threshold is not None:
        groups = _conservative_split_groups(similarity_matrix, groups, threshold)
    
    for group_idx, group in enumerate(groups):
        group_size = len(group)
        
        # Calculate pairwise similarities within the group
        group_similarities = []
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                sim = similarity_matrix[group[i]][group[j]]
                group_similarities.append(sim)
        
        avg_similarity = np.mean(group_similarities) if group_similarities else 0
        min_similarity = np.min(group_similarities) if group_similarities else 0
        
        # Find representative product (highest average similarity to others)
        representative_idx = _find_representative(similarity_matrix, group)
        
        # Prepare result
        analysis = {
            'group_id': f'G{group_idx + 1:03d}',
            'member_indices': group,
            'group_size': group_size,
            'avg_similarity': avg_similarity,
            'min_similarity': min_similarity,
            'representative_idx': representative_idx,
            'pairwise_similarities': group_similarities
        }
        
        # Add product names if provided
        if product_names:
            analysis['member_names'] = [product_names[i] for i in group]
            analysis['representative_name'] = product_names[representative_idx]
        
        analyses.append(analysis)
    
    return analyses


def _conservative_split_groups(similarity_matrix: np.ndarray,
                               groups: List[List[int]],
                               threshold: float) -> List[List[int]]:
    """Split connected components conservatively using representative affinity."""
    refined_groups = []

    for group in groups:
        if len(group) <= 2:
            refined_groups.append(group)
            continue

        remaining = set(group)
        while remaining:
            if len(remaining) == 1:
                break

            candidates = list(remaining)
            rep_idx = max(
                candidates,
                key=lambda i: np.mean([
                    similarity_matrix[i][j] for j in candidates if j != i
                ]) if len(candidates) > 1 else 0,
            )

            cluster = [
                idx for idx in candidates
                if idx == rep_idx or similarity_matrix[rep_idx][idx] >= threshold
            ]

            if len(cluster) > 1:
                refined_groups.append(sorted(cluster))

            remaining -= set(cluster)

    return refined_groups


def compute_group_evolution(similarity_matrix: np.ndarray,
                           product_names: List[str],
                           thresholds: List[int],
                           min_group_size: int = 2) -> pd.DataFrame:
    """
    Compute how groups evolve across different similarity thresholds.
    
    Groups are established at the lowest threshold and then tracked
    as members drop out at higher thresholds.
    
    Args:
        similarity_matrix: NxN matrix of similarity scores
        product_names: List of product names
        thresholds: List of thresholds to analyze (ascending)
        min_group_size: Minimum group size to track
        
    Returns:
        DataFrame with evolution tracking data
    """
    # Sort thresholds to ensure ascending order
    thresholds = sorted(thresholds)
    lowest_threshold = thresholds[0]
    
    # Find groups at the lowest threshold
    base_groups = find_product_groups(similarity_matrix, lowest_threshold, method='union_find')
    
    # Filter by minimum size
    base_groups = [g for g in base_groups if len(g) >= min_group_size]
    
    # Find representatives for each group
    group_representatives = {}
    for i, group in enumerate(base_groups):
        rep_idx = _find_representative(similarity_matrix, group)
        group_representatives[f'G{i+1:03d}'] = {
            'members': set(group),
            'representative': rep_idx,
            'representative_name': product_names[rep_idx]
        }
    
    # Track evolution across thresholds
    evolution_data = []
    
    for threshold in thresholds:
        for group_id, group_info in group_representatives.items():
            rep_idx = group_info['representative']
            original_members = group_info['members']
            
            # Check which members are still in the group at this threshold
            current_members = []
            similarities = []
            
            for member_idx in original_members:
                # Member stays if similarity to representative >= threshold
                if similarity_matrix[rep_idx][member_idx] >= threshold:
                    current_members.append(member_idx)
                    similarities.append(similarity_matrix[rep_idx][member_idx])
            
            # Only include if group still meets minimum size
            if len(current_members) >= min_group_size:
                min_similarity = min(similarities) if similarities else 0
                avg_similarity = sum(similarities) / len(similarities) if similarities else 0
                
                # Add row for each current member
                for member_idx in current_members:
                    evolution_data.append({
                        'Group ID': group_id,
                        'Group Summary': group_info['representative_name'],
                        'Threshold': threshold,
                        'Product Name': product_names[member_idx],
                        'Is Representative': member_idx == rep_idx,
                        'Similarity to Representative': similarity_matrix[rep_idx][member_idx],
                        'Group Size at Threshold': len(current_members),
                        'Group Min Similarity': round(min_similarity, 2),
                        'Group Avg Similarity': round(avg_similarity, 2),
                        'In Group': True
                })
                
                # Add rows for members that dropped out (for comparison)
                dropped_members = original_members - set(current_members)
                for member_idx in dropped_members:
                    evolution_data.append({
                        'Group ID': group_id,
                        'Group Summary': group_info['representative_name'],
                        'Threshold': threshold,
                        'Product Name': product_names[member_idx],
                        'Is Representative': member_idx == rep_idx,
                        'Similarity to Representative': similarity_matrix[rep_idx][member_idx],
                        'Group Size at Threshold': len(current_members),
                        'Group Min Similarity': round(min_similarity, 2),
                        'Group Avg Similarity': round(avg_similarity, 2),
                        'In Group': False
                    })
    
    return pd.DataFrame(evolution_data)


def _find_representative(similarity_matrix: np.ndarray, group: List[int]) -> int:
    """
    Find the most representative product in a group based on centrality.
    
    Args:
        similarity_matrix: NxN matrix of similarity scores
        group: List of product indices in the group
        
    Returns:
        Index of the most representative product
    """
    if len(group) == 1:
        return group[0]
    
    # Calculate average similarity of each product to all others in the group
    avg_similarities = []
    for i in group:
        similarities = [similarity_matrix[i][j] for j in group if i != j]
        avg_similarities.append(np.mean(similarities) if similarities else 0)
    
    # Return the product with highest average similarity
    return group[np.argmax(avg_similarities)]


def create_grouped_results(analyses: List[Dict[str, Any]], 
                          product_data: pd.DataFrame,
                          display_columns: List[str] = None) -> pd.DataFrame:
    """
    Create a DataFrame with grouped results for display.
    
    Args:
        analyses: List of group analyses
        product_data: DataFrame with product information
        display_columns: Columns to include in the output
        
    Returns:
        DataFrame with grouped results
    """
    results = []
    
    for analysis in analyses:
        group_id = analysis['group_id']
        member_indices = analysis['member_indices']
        representative_idx = analysis['representative_idx']
        representative_name = analysis.get('representative_name', str(representative_idx))

        for i, member_idx in enumerate(member_indices):
            member_name = analysis.get('member_names', [str(idx) for idx in member_indices])[i]
            member_info = {
                'Group ID': group_id,
                'Group Summary': representative_name,
                'Product Name': member_name,
                'Is Representative': member_idx == representative_idx,
                'Group Size': analysis['group_size'],
                'Group Avg Similarity': round(float(analysis['avg_similarity']), 2),
                'Group Min Similarity': round(float(analysis['min_similarity']), 2),
            }

            if display_columns:
                for col in display_columns:
                    if col in product_data.columns:
                        member_info[col] = _to_serializable_scalar(product_data.iloc[member_idx][col])

            results.append(member_info)
    
    return pd.DataFrame(results)


def export_groups_flat(analyses: List[Dict[str, Any]], 
                      product_data: pd.DataFrame,
                      display_columns: List[str] = None) -> pd.DataFrame:
    """
    Export groups in flat format with group IDs.
    
    Args:
        analyses: List of group analyses
        product_data: DataFrame with product information
        display_columns: Columns to include in the output
        
    Returns:
        DataFrame with flat group export
    """
    results = []
    
    for analysis in analyses:
        group_id = analysis['group_id']
        member_indices = analysis['member_indices']
        
        representative_name = analysis.get('representative_name', str(analysis['representative_idx']))

        for i, member_idx in enumerate(member_indices):
            row = {
                'Group ID': group_id,
                'Group Summary': representative_name,
                'Product Name': analysis.get('member_names', [str(member_idx)])[i],
                'Is Representative': member_idx == analysis['representative_idx'],
                'Group Size': analysis['group_size'],
                'Group Avg Similarity': round(float(analysis['avg_similarity']), 2),
                'Group Min Similarity': round(float(analysis['min_similarity']), 2),
            }
            
            # Add additional columns if specified
            if display_columns:
                for col in display_columns:
                    if col in product_data.columns:
                        row[col] = _to_serializable_scalar(product_data.iloc[member_idx][col])
            
            results.append(row)
    
    return pd.DataFrame(results)


def filter_groups(analyses: List[Dict[str, Any]], 
                 min_group_size: int = 2,
                 max_groups: int = None,
                 sort_by: str = 'size') -> List[Dict[str, Any]]:
    """
    Filter and sort groups based on criteria.
    
    Args:
        analyses: List of group analyses
        min_group_size: Minimum group size to include
        max_groups: Maximum number of groups to return
        sort_by: Sort criteria ('size', 'avg_similarity', 'min_similarity')
        
    Returns:
        Filtered and sorted list of analyses
    """
    # Filter by minimum size
    filtered = [a for a in analyses if a['group_size'] >= min_group_size]
    
    # Sort
    if sort_by == 'size':
        filtered.sort(key=lambda x: x['group_size'], reverse=True)
    elif sort_by == 'avg_similarity':
        filtered.sort(key=lambda x: x['avg_similarity'], reverse=True)
    elif sort_by == 'min_similarity':
        filtered.sort(key=lambda x: x['min_similarity'], reverse=True)
    
    # Limit number of groups
    if max_groups:
        filtered = filtered[:max_groups]
    
    return filtered


def get_group_analyses(similarity_matrix: np.ndarray,
                      product_names: List[str],
                      similarity_threshold: float,
                      min_group_size: int = 2,
                      max_groups: int = None,
                      conservative_grouping: bool = True) -> List[Dict[str, Any]]:
    """Build filtered group analyses from a similarity matrix using consistent settings."""
    groups = find_product_groups(similarity_matrix, threshold=similarity_threshold)
    if not groups:
        return []

    analyses = analyze_groups(
        similarity_matrix,
        groups,
        product_names,
        threshold=similarity_threshold,
        conservative=conservative_grouping,
    )

    return filter_groups(
        analyses,
        min_group_size=min_group_size,
        max_groups=max_groups,
        sort_by='size',
    )
