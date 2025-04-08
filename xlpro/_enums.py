class FunctionTypes:
    array_or_value = 0
    py_object = 1

    @classmethod
    def as_list(cls):
        return [value for key, value in vars(cls).items() if isinstance(value, int)]
    

