class FunctionTypes:
    default = 0
    figure = 1
    # jsonified = 2

    @classmethod
    def as_list(cls):
        return [value for key, value in vars(cls).items() if isinstance(value, int)]
    

