"use client";

import { Skeleton } from "./ui/skeleton";

export default function Loading() {
  return (
    <>
      <div className="flex flex-col gap-4">
        <div className="flex items-center gap-4">
          <Skeleton className="h-12 w-48" />
          <Skeleton className="h-12 w-24" />
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
      </div>
      <div className="flex flex-col gap-4">
        <div className="flex items-center gap-4">
          <Skeleton className="h-12 w-48" />
          <Skeleton className="h-12 w-24" />
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
        <div>
          <Skeleton className="h-12 w-76">
            <Skeleton className="h-12 w-24" />
          </Skeleton>
        </div>
      </div>
    </>
  );
}
