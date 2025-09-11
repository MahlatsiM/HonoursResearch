#!/bin/bash
#SBATCH --job-name=Cars  # Job name
#SBATCH --error=error.log                      # Standard error log file
#SBATCH --output=output.log                    # Standard output log file
#SBATCH --ntasks=1                             # Number of tasks
#SBATCH --cpus-per-task=4                      # CPU cores per task
#SBATCH --mem=16G                              # Increased memory for GPU workloads
#SBATCH --gres=gpu:1                           # Request 1 GPU
#SBATCH --partition=GPU                        # GPU partition (adjust based on cluster)
#SBATCH --mail-user=202215639@spu.ac.za
#SBATCH --mail-type=BEGIN,END,FAIL,TIME_LIMIT_80
#SBATCH --time=00:10:00                        # Max runtime (hh:mm:ss)

# Load required modules
module load python/3.10  
module load cuda/11.8   # Load CUDA (adjust version based on cluster)
module load cudnn/8.2   # Load cuDNN (adjust version based on cluster)

# Install required Python libraries
pip install --user pandas numpy seaborn matplotlib scikit-learn imbalanced-learn torch tensorflow-gpu

# Run the Python script
python Cars_Model_Building.ipynb
