package com.clmgni.api.utils

class Util {
    companion object {
        fun removeSpecialCaracterFromString(obj : String) : String {
            return obj.replace("[^a-zA-Z0-9]".toRegex(), "")
        }
    }
}