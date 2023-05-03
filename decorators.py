def singleton(cls: type):
    instance_dict = {}

    def get_instance(*args, **kwargs):
        if cls not in instance_dict:
            instance_dict[cls] = cls(*args, **kwargs)
        return instance_dict[cls]

    return get_instance
