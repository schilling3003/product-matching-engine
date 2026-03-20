import numpy as np
import pandas as pd
from typing import List, Tuple, Dict, Any

try:
    import networkx as nx
    HAS_NETWORKX = True
except ImportError:
    HAS_NETWORKX = False


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
                  product_names: List[str] = None) -> List[Dict[str, Any]]:
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
        
        # Add summary row for the group
        summary = {
            'Group ID': group_id,
            'Product Name': f"[Group Summary] {analysis.get('representative_name', 'Representative')}",
            'Role': 'Representative',
            'Group Size': analysis['group_size'],
            'Avg Similarity': f"{analysis['avg_similarity']:.2f}%",
            'Min Similarity': f"{analysis['min_similarity']:.2f}%"
        }
        
        # Add additional columns if specified
        if display_columns:
            rep_idx = analysis['representative_idx']
            for col in display_columns:
                if col in product_data.columns:
                    summary[col] = product_data.iloc[rep_idx][col]
        
        results.append(summary)
        
        # Add member details
        for i, member_idx in enumerate(member_indices):
            if member_idx == analysis['representative_idx']:
                continue  # Skip representative as it's already shown
                
            member_info = {
                'Group ID': group_id,
                'Product Name': analysis.get('member_names', [str(member_idx)])[i],
                'Role': 'Member',
                'Group Size': '',
                'Avg Similarity': '',
                'Min Similarity': ''
            }
            
            # Add additional columns if specified
            if display_columns:
                for col in display_columns:
                    if col in product_data.columns:
                        member_info[col] = product_data.iloc[member_idx][col]
            
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
        
        for i, member_idx in enumerate(member_indices):
            row = {
                'Group ID': group_id,
                'Product Name': analysis.get('member_names', [str(member_idx)])[i],
                'Is Representative': member_idx == analysis['representative_idx'],
                'Group Size': analysis['group_size'],
                'Group Avg Similarity': f"{analysis['avg_similarity']:.2f}%"
            }
            
            # Add additional columns if specified
            if display_columns:
                for col in display_columns:
                    if col in product_data.columns:
                        row[col] = product_data.iloc[member_idx][col]
            
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
