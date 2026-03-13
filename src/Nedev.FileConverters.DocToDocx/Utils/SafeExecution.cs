using System;
using System.Collections.Generic;
using System.IO;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Provides safe execution wrappers for error-prone operations with recovery capabilities.
/// </summary>
public static class SafeExecution
{
    /// <summary>
    /// Executes an action with automatic retry on transient failures.
    /// </summary>
    public static void ExecuteWithRetry(
        Action operation,
        int maxRetries = 3,
        TimeSpan? delayBetweenRetries = null,
        Func<Exception, bool>? isRetryable = null)
    {
        isRetryable ??= IsTransientError;
        delayBetweenRetries ??= TimeSpan.FromMilliseconds(100);

        var lastException = default(Exception);

        for (int attempt = 0; attempt < maxRetries; attempt++)
        {
            try
            {
                operation();
                return;
            }
            catch (Exception ex) when (isRetryable(ex) && attempt < maxRetries - 1)
            {
                lastException = ex;
                Logger.Warning($"Operation failed (attempt {attempt + 1}/{maxRetries}), retrying...", ex);
                System.Threading.Thread.Sleep(delayBetweenRetries.Value);
            }
        }

        if (lastException != null)
        {
            throw new InvalidOperationException(
                $"Operation failed after {maxRetries} attempts. Last error: {lastException.Message}",
                lastException);
        }
    }

    /// <summary>
    /// Executes a function with automatic retry on transient failures.
    /// </summary>
    public static T ExecuteWithRetry<T>(
        Func<T> operation,
        int maxRetries = 3,
        TimeSpan? delayBetweenRetries = null,
        Func<Exception, bool>? isRetryable = null)
    {
        isRetryable ??= IsTransientError;
        delayBetweenRetries ??= TimeSpan.FromMilliseconds(100);

        var lastException = default(Exception);

        for (int attempt = 0; attempt < maxRetries; attempt++)
        {
            try
            {
                return operation();
            }
            catch (Exception ex) when (isRetryable(ex) && attempt < maxRetries - 1)
            {
                lastException = ex;
                Logger.Warning($"Operation failed (attempt {attempt + 1}/{maxRetries}), retrying...", ex);
                System.Threading.Thread.Sleep(delayBetweenRetries.Value);
            }
        }

        if (lastException != null)
        {
            throw new InvalidOperationException(
                $"Operation failed after {maxRetries} attempts. Last error: {lastException.Message}",
                lastException);
        }

        throw new InvalidOperationException("Operation failed with no exception captured");
    }

    /// <summary>
    /// Executes an operation and returns a Result, catching any exceptions.
    /// </summary>
    public static Result<T> ExecuteSafe<T>(Func<T> operation, string context)
    {
        try
        {
            var result = operation();
            return Result<T>.Success(result);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in {context}", ex);
            return Result<T>.Failure($"{context}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Executes an action and returns a Result, catching any exceptions.
    /// </summary>
    public static Result ExecuteSafe(Action operation, string context)
    {
        try
        {
            operation();
            return Result.Success();
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in {context}", ex);
            return Result.Failure($"{context}: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Executes an operation with a fallback value on failure.
    /// </summary>
    public static T ExecuteWithFallback<T>(Func<T> operation, T fallbackValue, string context)
    {
        try
        {
            return operation();
        }
        catch (Exception ex)
        {
            Logger.Warning($"Using fallback value in {context}: {ex.Message}", ex);
            return fallbackValue;
        }
    }

    /// <summary>
    /// Executes an operation and ignores any non-critical exceptions.
    /// </summary>
    public static void ExecuteBestEffort(Action operation, string context)
    {
        try
        {
            operation();
        }
        catch (Exception ex) when (!IsCriticalError(ex))
        {
            Logger.Warning($"Best-effort operation failed in {context}, continuing...", ex);
        }
    }

    /// <summary>
    /// Determines if an exception is transient and retryable.
    /// </summary>
    private static bool IsTransientError(Exception ex)
    {
        return ex is IOException or
               UnauthorizedAccessException or
               TimeoutException;
    }

    /// <summary>
    /// Determines if an exception is critical and should not be ignored.
    /// </summary>
    private static bool IsCriticalError(Exception ex)
    {
        return ex is OutOfMemoryException or
               StackOverflowException or
               AccessViolationException;
    }
}

/// <summary>
/// Extension methods for safe execution patterns.
/// </summary>
public static class SafeExecutionExtensions
{
    /// <summary>
    /// Safely disposes an object, ignoring any exceptions.
    /// </summary>
    public static void SafeDispose(this IDisposable? disposable)
    {
        if (disposable == null) return;

        try
        {
            disposable.Dispose();
        }
        catch (Exception ex)
        {
            Logger.Debug($"Exception during disposal of {disposable.GetType().Name}: {ex.Message}");
        }
    }

    /// <summary>
    /// Safely disposes all items in a collection.
    /// </summary>
    public static void SafeDisposeAll<T>(this IEnumerable<T>? items) where T : IDisposable
    {
        if (items == null) return;

        foreach (var item in items)
        {
            item?.SafeDispose();
        }
    }

    /// <summary>
    /// Executes an action if the object is not null.
    /// </summary>
    public static void IfNotNull<T>(this T? obj, Action<T> action) where T : class
    {
        if (obj != null)
        {
            action(obj);
        }
    }

    /// <summary>
    /// Executes a function if the object is not null, returning a default value otherwise.
    /// </summary>
    public static TResult IfNotNull<T, TResult>(this T? obj, Func<T, TResult> func, TResult defaultValue) where T : class
    {
        return obj != null ? func(obj) : defaultValue;
    }
}
