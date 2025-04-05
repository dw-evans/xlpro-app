class FunctionTypes:
    array_or_value = 0
    figure = 1
    py_object = 2

    @classmethod
    def as_list(cls):
        return [value for key, value in vars(cls).items() if isinstance(value, int)]
    

