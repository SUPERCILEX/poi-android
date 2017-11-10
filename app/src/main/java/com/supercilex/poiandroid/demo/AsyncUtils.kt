package com.supercilex.poiandroid.demo

import com.google.android.gms.tasks.Task
import com.google.android.gms.tasks.Tasks
import java.util.concurrent.Callable
import java.util.concurrent.Executor
import java.util.concurrent.Executors

inline fun <T> async(crossinline block: () -> T): Task<T> =
        AsyncTaskExecutor.execute(Callable { block() })

object AsyncTaskExecutor : Executor {
    private val service = Executors.newCachedThreadPool()

    fun <TResult> execute(callable: Callable<TResult>): Task<TResult> = Tasks.call(this, callable)

    override fun execute(runnable: Runnable) {
        service.submit(runnable)
    }
}
