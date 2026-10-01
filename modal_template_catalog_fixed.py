"""Modal runner for a private, cold-start-friendly Qwen2.5-VL-7B server.
Place beside template_catalog.py. Deploy server once, then call catalog independently.
"""
import json
import os
from pathlib import Path
import modal

app=modal.App("rao-template-catalog")
MODEL="Qwen/Qwen2.5-VL-7B-Instruct"
cache=modal.Volume.from_name("rao-hf-cache",create_if_missing=True)
cpu_image=(modal.Image.debian_slim(python_version="3.11")
 .apt_install("libreoffice","poppler-utils","fonts-liberation")
 .pip_install("python-pptx==1.0.2","openai>=1,<3","pdf2image","pillow")
 .add_local_file("template_catalog.py","/root/template_catalog.py"))
# Model container has compiler toolchain, though eager startup does not require TorchInductor JIT.
gpu_image=(modal.Image.from_registry("nvidia/cuda:12.9.0-devel-ubuntu22.04",add_python="3.12")
 .entrypoint([]).uv_pip_install("vllm==0.21.0"))

@app.function(image=cpu_image,cpu=2,memory=4096,timeout=1800,)
def catalog(pptx:bytes,filename:str,mode:str,base_url:str|None,model:str|None,max_calls:int=2):
 import sys,tempfile
 sys.path.insert(0,"/root")
 from template_catalog import run
 with tempfile.TemporaryDirectory() as temp:
  path=Path(temp)/Path(filename).name
  path.write_bytes(pptx)
  return run(str(path),mode,base_url,model,os.getenv("LLM_API_KEY","EMPTY"),max_calls)

@app.function(image=gpu_image,gpu="A100",volumes={"/root/.cache/huggingface":cache},
              timeout=1800,scaledown_window=600)
@modal.web_server(port=8000,startup_timeout=1200)
def vision_server():
 import subprocess
 cmd=["vllm","serve",MODEL,"--served-model-name",MODEL,"--host","0.0.0.0","--port","8000",
      "--dtype","bfloat16","--max-model-len","4096","--max-num-seqs","2",
      "--max-num-batched-tokens","4096","--gpu-memory-utilization","0.8",
      "--limit-mm-per-prompt",'{"image":1}',"--enforce-eager"]
 subprocess.Popen(cmd,env={**os.environ,"VLLM_USE_FLASHINFER_SAMPLER":"0"})

@app.local_entrypoint()
def main(pptx:str,mode:str="heuristic",output:str="template_catalog.json",
         base_url:str="",model:str="",max_calls:int=2):
 if mode not in ("heuristic","text","vlm"):raise ValueError("mode must be heuristic,text,vlm")
 if mode!="heuristic" and not base_url:
  base_url=vision_server.get_web_url()+"/v1"
  model=MODEL
 if mode!="heuristic" and not model:raise ValueError("model ID is required")
 path=Path(pptx)
 data=catalog.remote(path.read_bytes(),path.name,mode,base_url or None,model or None,max_calls)
 Path(output).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
 print(f"Wrote {output}; model_errors={sum('model_error' in s for s in data['slides'])}")
