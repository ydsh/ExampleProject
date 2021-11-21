package com.example.excel.util;

import java.util.function.Supplier;

public final class FuncUtil {
	private FuncUtil() {
	}

	public static <T> T create(Supplier<T> supplier) {
		return supplier.get();
	}
}
