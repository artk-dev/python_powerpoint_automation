networkinterfaces = {"eni-04855e4759dbf2eb8":"vpce-09a7202c8c55786a0", "eni-0789aa367acb505b8":"VPC Endpoint Interface vpce-09eb95d9d49b9eb67", "eni-093063ee4f594ceb4":"", "eni-0a4c3b997f89ce13c": eni-0aa4e6e356c89f753 eni-0ca0be4d80b36bbfb eni-0ce1eec12d80a6678}

for key in networkinterfaces.keys():
    print(f'Network Interface {key} connects to {networkinterfaces[key]}')
    