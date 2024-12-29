
"""
Callback structure

write async results to a register
when a function completes, it has data which signals which cells (callers) need to be recomputed

# when a function is called, the output is given a unique id in the cached results.
# the cache is filled with the temporary placeholder results until it is complete
# once the uid has been flagged complete, it looks up
# the uid can cache a set of arguments func(a, b, c) so we can look up any existing ones in the past (potentially)
# This caching of results would require a cache clear option to the formulas, maybe with an optional defaulted to no
# But at the same time we might want to prevent any big calculations from being executed unknowingly 

# Is there an opportunity to have some helper methods that can see which cells are calling the xlpro functions

# Oh and by the way we need to nest everything down a level so each workbook has its own memory space if we are doing
# the persistent memory option

# How can we issue the callback
# if a function knows its caller, it can issue a Range.Recalculate callback pretty easily.

# All of this relies on recalculate doing a cache lookup first. and if it misses, deferring the calculation to a background
# python process

# maps the function uuid to the live result
results_register = {
    # "uuid": <T>result,
    12321453:
}

# look up the function to the uuid
function_register = {
    hash(func): uuid
}

caller_register = {
    workbook.sheet.a1: uuid
}

# an async filler function can be called in the meantime for each life function
# filler function
lambda repr, t: f"{repr} has been executing for {t} seconds"

# as these are seen as complete
completion_register = {
    uuid: True/False
}

# when a function is marked complete the called must be ordered to recalculate and points to
# a cached output of the function 

"""
