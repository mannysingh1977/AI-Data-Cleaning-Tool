import os
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

def load_text_file(file_path):
    """Load text from a file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def compare_documents(file1_path, file2_path, model):
    """Compare two documents and return similarity score"""
    
    # Load text from both files
    print(f"\n📄 Loading documents...")
    text1 = load_text_file(file1_path)
    text2 = load_text_file(file2_path)
    
    print(f"   Document 1: {len(text1)} characters")
    print(f"   Document 2: {len(text2)} characters")
    
    # Generate embeddings
    print(f"\n🧠 Generating embeddings...")
    embedding1 = model.encode(text1)
    embedding2 = model.encode(text2)
    
    print(f"   Embedding dimensions: {len(embedding1)}")
    
    # Calculate similarity
    print(f"\n🔍 Calculating similarity...")
    similarity = cosine_similarity(
        [embedding1], 
        [embedding2]
    )[0][0]
    
    return similarity, text1, text2

if __name__ == "__main__":
    print("="*70)
    print("DOCUMENT SIMILARITY COMPARISON")
    print("="*70)
    
    # Load the embedding model (this happens once)
    print("\n⚙️  Loading AI model (this may take a moment first time)...")
    model = SentenceTransformer('all-MiniLM-L6-v2')  # Small, fast model
    print("   ✓ Model loaded!")
    
    # Folder with extracted text
    extracted_folder = "extracted_text"
    
    # List available files
    print(f"\n📁 Available files in '{extracted_folder}/':")
    files = [f for f in os.listdir(extracted_folder) if f.endswith('.txt')]
    for i, file in enumerate(files, 1):
        print(f"   {i}. {file}")
    
    if len(files) < 2:
        print("\n❌ Error: You need at least 2 .txt files to compare!")
        print(f"   Currently you have {len(files)} file(s).")
        exit()
    
    # Ask user to pick 2 files
    print("\n" + "="*70)
    file1_num = int(input("Enter number of FIRST document: ")) - 1
    file2_num = int(input("Enter number of SECOND document: ")) - 1
    
    file1_path = os.path.join(extracted_folder, files[file1_num])
    file2_path = os.path.join(extracted_folder, files[file2_num])
    
    # Compare them
    print("\n" + "="*70)
    similarity, text1, text2 = compare_documents(file1_path, file2_path, model)
    
    # Show results
    print("\n" + "="*70)
    print("RESULTS")
    print("="*70)
    print(f"\n📊 Similarity Score: {similarity:.4f}")
    print(f"   (Range: -1.0 = opposite, 0.0 = unrelated, 1.0 = identical)")

    # Interpretation
    if similarity >= 0.95:
        print("   🔴 VERY HIGH - Likely duplicates or near-duplicates")
    elif similarity >= 0.80:
        print("   🟠 HIGH - Very similar content, possibly versions")
    elif similarity >= 0.60:
        print("   🟡 MODERATE - Related topics or overlapping content")
    elif similarity >= 0.40:
        print("   🟢 LOW - Some similarities but different content")
    else:
        print("   ⚪ VERY LOW - Completely different documents")
        
    # Show text previews
    print(f"\n📄 Document 1 preview:")
    print(f"   {text1[:200].replace(chr(10), ' ')}...")
    
    print(f"\n📄 Document 2 preview:")
    print(f"   {text2[:200].replace(chr(10), ' ')}...")
    
    print("\n" + "="*70)
