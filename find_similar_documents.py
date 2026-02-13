import os
import json
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

# Map source folders to file types
SOURCE_FOLDERS = {
    "word_files": "docx",
    "pdf_files": "pdf",
    "pptx_files": "pptx",
    "xlsx_files": "xlsx",
}

def get_source_type(filename):
    """Determine the original file type by checking which source folder has a matching file"""
    base_name = os.path.splitext(filename)[0]
    for folder, file_type in SOURCE_FOLDERS.items():
        if os.path.exists(folder):
            for f in os.listdir(folder):
                if os.path.splitext(f)[0] == base_name:
                    return file_type
    return "unknown"

def load_all_documents(folder_path):
    """Load all .txt files from folder with their source type"""
    documents = {}

    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
                source_type = get_source_type(filename)
                documents[filename] = {"text": text, "type": source_type}
                print(f"📄 Loaded {filename} ({len(text)} chars) [{source_type}]")

    return documents

def chunk_text(text, chunk_size=500, overlap=100):
    """Split text into overlapping chunks of approximately chunk_size words"""
    words = text.split()
    if len(words) <= chunk_size:
        return [text]

    chunks = []
    start = 0
    while start < len(words):
        end = start + chunk_size
        chunk = " ".join(words[start:end])
        chunks.append(chunk)
        start += chunk_size - overlap
    return chunks

def find_similar_pairs(doc_embeddings, documents, threshold=0.6, mode="all", selected_types=None):
    """Find all document pairs above similarity threshold using chunked embeddings"""
    similar_pairs = []
    doc_names = list(documents.keys())

    for i in range(len(doc_names)):
        for j in range(i+1, len(doc_names)):
            name_i, name_j = doc_names[i], doc_names[j]
            type_i = documents[name_i]["type"]
            type_j = documents[name_j]["type"]

            # Filter based on comparison mode
            if mode == "same_type" and type_i != type_j:
                continue
            if mode == "select" and selected_types:
                if type_i not in selected_types or type_j not in selected_types:
                    continue
                if type_i == type_j:
                    continue

            # Compare all chunk pairs, take the max similarity
            sim_matrix = cosine_similarity(doc_embeddings[name_i], doc_embeddings[name_j])
            max_sim = float(np.max(sim_matrix))

            if max_sim >= threshold:
                similar_pairs.append({
                    'doc1': name_i,
                    'doc2': name_j,
                    'type1': type_i,
                    'type2': type_j,
                    'similarity': max_sim,
                    'preview1': documents[name_i]["text"][:100],
                    'preview2': documents[name_j]["text"][:100]
                })

    return similar_pairs

def group_similar_documents(similar_pairs):
    """Group documents into clusters using Union-Find based on similar pairs"""
    # Union-Find: each document starts as its own group
    parent = {}

    def find(x):
        if parent[x] != x:
            parent[x] = find(parent[x])
        return parent[x]

    def union(a, b):
        root_a, root_b = find(a), find(b)
        if root_a != root_b:
            parent[root_b] = root_a

    # Initialize each document as its own parent
    all_docs = set()
    for pair in similar_pairs:
        all_docs.add(pair['doc1'])
        all_docs.add(pair['doc2'])
    for doc in all_docs:
        parent[doc] = doc

    # Merge documents that are similar
    for pair in similar_pairs:
        union(pair['doc1'], pair['doc2'])

    # Build clusters
    clusters = {}
    for doc in all_docs:
        root = find(doc)
        if root not in clusters:
            clusters[root] = []
        clusters[root].append(doc)

    # Build a lookup of max similarity between each pair in the cluster
    pair_lookup = {}
    for pair in similar_pairs:
        key = tuple(sorted([pair['doc1'], pair['doc2']]))
        pair_lookup[key] = pair['similarity']

    # Format clusters with info
    result = []
    for root, members in clusters.items():
        if len(members) < 2:
            continue
        # Calculate average similarity within the cluster
        sims = []
        for i in range(len(members)):
            for j in range(i+1, len(members)):
                key = tuple(sorted([members[i], members[j]]))
                if key in pair_lookup:
                    sims.append(pair_lookup[key])
        avg_sim = sum(sims) / len(sims) if sims else 0
        max_sim = max(sims) if sims else 0

        result.append({
            'members': sorted(members),
            'size': len(members),
            'avg_similarity': avg_sim,
            'max_similarity': max_sim,
        })

    # Sort by size (largest first), then by avg similarity
    result.sort(key=lambda x: (x['size'], x['avg_similarity']), reverse=True)
    return result

def main():
    print("="*80)
    print("FIND SIMILAR DOCUMENTS (Duplicate Detection Prototype)")
    print("="*80)

    # Folder with extracted text
    extracted_folder = "extracted_text"

    if not os.path.exists(extracted_folder):
        print(f"❌ Folder '{extracted_folder}' not found!")
        print("Run your extraction script first.")
        return

    # Load metadata if available
    metadata_folder = "metadata"
    results_folder = "results"
    os.makedirs(results_folder, exist_ok=True)

    metadata_path = os.path.join(metadata_folder, "metadata.json")
    metadata = {}
    if os.path.exists(metadata_path):
        with open(metadata_path, 'r', encoding='utf-8') as f:
            metadata = json.load(f)
        print(f"📋 Loaded metadata for {len(metadata)} documents")
    else:
        print("⚠️  No metadata.json found. Run extract.py first for metadata support.")

    # Load all documents
    documents = load_all_documents(extracted_folder)

    if len(documents) < 2:
        print("❌ Need at least 2 documents to compare!")
        return

    # Show available types
    available_types = sorted(set(doc["type"] for doc in documents.values()))
    type_counts = {}
    for doc in documents.values():
        type_counts[doc["type"]] = type_counts.get(doc["type"], 0) + 1

    print(f"\n📁 Found document types:")
    for t in available_types:
        print(f"   - {t} ({type_counts[t]} files)")

    # Ask user for comparison mode
    print(f"\n🔧 Comparison mode:")
    print(f"   [1] Compare ALL documents (cross-format, finds docx vs pdf duplicates)")
    print(f"   [2] Compare within SAME TYPE only (docx vs docx, pdf vs pdf, etc.)")
    print(f"   [3] Select SPECIFIC types to compare")

    choice = input("\n   Choose mode (1/2/3): ").strip()

    mode = "all"
    selected_types = None

    if choice == "2":
        mode = "same_type"
        print("   → Comparing within same file type only")
    elif choice == "3":
        print(f"\n   Available types: {', '.join(available_types)}")
        type_input = input("   Enter types to compare (comma-separated, e.g. docx,pdf / docx,xlsx / pdf,xlsx): ").strip()
        selected_types = [t.strip() for t in type_input.split(",")]
        mode = "select"
        print(f"   → Comparing only: {', '.join(selected_types)}")
    else:
        print("   → Comparing all documents across all types")

    # Load model
    print("\n⚙️  Loading embedding model...")
    model = SentenceTransformer('all-MiniLM-L6-v2')

    # Chunk and embed all documents
    print("\n🧠 Chunking and generating embeddings...")
    doc_names = list(documents.keys())
    doc_chunks = {}
    doc_embeddings = {}
    total_chunks = 0

    for name in doc_names:
        chunks = chunk_text(documents[name]["text"])
        doc_chunks[name] = chunks
        doc_embeddings[name] = model.encode(chunks)
        total_chunks += len(chunks)

    print(f"   ✓ {total_chunks} chunks from {len(doc_names)} documents (384 dims each)")

    # Find similar pairs using max chunk similarity
    threshold = 0.6
    print(f"\n🎯 Looking for pairs with similarity >= {threshold:.3f}...")

    similar_pairs = find_similar_pairs(doc_embeddings, documents, threshold, mode, selected_types)

    # Sort by similarity (highest first)
    similar_pairs.sort(key=lambda x: x['similarity'], reverse=True)

    # Show results
    print("\n" + "="*80)
    print("SIMILAR DOCUMENT PAIRS")
    print("="*80)

    if similar_pairs:
        for pair in similar_pairs:
            print(f"\n🔗 SIMILARITY: {pair['similarity']:.4f}")

            # Show doc1 with metadata
            meta1 = metadata.get(pair['doc1'], {})
            print(f"   📄 {pair['doc1']} [{pair['type1']}]")
            if meta1:
                print(f"      Author: {meta1.get('author', '?')} | Modified: {meta1.get('modified', '?')}")

            # Show doc2 with metadata
            meta2 = metadata.get(pair['doc2'], {})
            print(f"   📄 {pair['doc2']} [{pair['type2']}]")
            if meta2:
                print(f"      Author: {meta2.get('author', '?')} | Modified: {meta2.get('modified', '?')}")

            # Similarity level
            if pair['similarity'] >= 0.95:
                print("   🔴 VERY HIGH similarity: Exact duplicates → keep one, delete other")
            elif pair['similarity'] >= 0.85:
                print("   🔴 HIGH similarity: Very similar → likely versions → merge candidates")
            elif pair['similarity'] >= 0.7:
                print("   🟠 MEDIUM similarity: related content → review for overlap")
            else:
                print("   🟡 LOW similarity → loosely related → consider grouping")

            # Cross-format warning
            if pair['type1'] != pair['type2']:
                print(f"   ⚠️  Cross-format match ({pair['type1']} vs {pair['type2']})")
    else:
        print("✅ No highly similar documents found (all below threshold)")

    # Group into clusters
    if similar_pairs:
        clusters = group_similar_documents(similar_pairs)

        print("\n" + "="*80)
        print("DOCUMENT CLUSTERS")
        print("="*80)

        for i, cluster in enumerate(clusters, 1):
            # Cluster level based on average similarity
            if cluster['avg_similarity'] >= 0.85:
                level = "🔴 HIGH"
            elif cluster['avg_similarity'] >= 0.7:
                level = "🟠 MEDIUM"
            else:
                level = "🟡 LOW"

            print(f"\n📁 Cluster {i} — {cluster['size']} documents — {level} similarity")
            print(f"   Avg: {cluster['avg_similarity']:.4f} | Max: {cluster['max_similarity']:.4f}")
            for member in cluster['members']:
                doc_type = documents[member]["type"]
                meta = metadata.get(member, {})
                author = meta.get('author', '?')
                modified = meta.get('modified', '?')
                print(f"   - {member} [{doc_type}] | Author: {author} | Modified: {modified}")

    # Export results to JSON
    results = {
        "settings": {
            "threshold": threshold,
            "mode": mode,
            "selected_types": selected_types,
            "total_documents": len(documents),
        },
        "pairs": [],
        "clusters": [],
    }

    for pair in similar_pairs:
        pair_data = {
            "doc1": pair['doc1'],
            "doc2": pair['doc2'],
            "type1": pair['type1'],
            "type2": pair['type2'],
            "similarity": round(pair['similarity'], 4),
            "level": "very_high" if pair['similarity'] >= 0.95
                     else "high" if pair['similarity'] >= 0.85
                     else "medium" if pair['similarity'] >= 0.7
                     else "low",
            "cross_format": pair['type1'] != pair['type2'],
            "doc1_metadata": metadata.get(pair['doc1'], {}),
            "doc2_metadata": metadata.get(pair['doc2'], {}),
        }
        results["pairs"].append(pair_data)

    if similar_pairs:
        for i, cluster in enumerate(clusters, 1):
            cluster_data = {
                "cluster_id": i,
                "size": cluster['size'],
                "avg_similarity": round(cluster['avg_similarity'], 4),
                "max_similarity": round(cluster['max_similarity'], 4),
                "level": "high" if cluster['avg_similarity'] >= 0.85
                         else "medium" if cluster['avg_similarity'] >= 0.7
                         else "low",
                "members": [],
            }
            for member in cluster['members']:
                cluster_data["members"].append({
                    "filename": member,
                    "type": documents[member]["type"],
                    "metadata": metadata.get(member, {}),
                })
            results["clusters"].append(cluster_data)

    results_path = os.path.join(results_folder, "results.json")
    with open(results_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, default=str)

    print(f"\n📊 Summary: {len(similar_pairs)} pairs found, {len(clusters) if similar_pairs else 0} clusters")
    print(f"💾 Results exported to {results_path}")
    print("="*80)

if __name__ == "__main__":
    main()
