DLSSP Pipeline module (dlssp_pipeline.py)

Files created:
- /mnt/data/dlssp_pipeline.py  (main module)
- /mnt/data/dlssp_example.xlsx (created when running example)

How to run locally:
1) Copy dlssp_pipeline.py into your VSCode project.
2) Install dependencies: pip install pandas networkx xlsxwriter openpyxl
3) Run Python interactive:
   >>> from dlssp_pipeline import run_dlssp_example, run_pipeline_from_excel
   >>> run_dlssp_example()   # creates an example input and runs the pipeline
   OR
   >>> run_pipeline_from_excel('/path/to/your_input.xlsx')

Input Excel format (sheets):
- Params: columns ('param','value')
- Lanes: columns ('lane_id','l_pos')
- Trays: columns ('k_pos','sku','wave' optional)
- Orders: columns ('order_id','class','due','weight' optional)
- Demand: columns ('order_id','sku','qty')
- Waves (optional): columns ('wave_id','release_time')

Notes:
- This implementation is a rebuild focused on clarity and DLSSP concepts (waves, release times, Ut, gridlock).
- ALNS is a simple remove-and-reinsert local search; you can tune 'alns_iters' param.
- It returns a CSV schedule and prints a summary.

I cannot push directly to your VSCode. The files are stored in the runtime at:
- /mnt/data/dlssp_pipeline.py
You can download them from the notebook environment or copy into your VSCode workspace.
